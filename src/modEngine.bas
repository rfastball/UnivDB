Attribute VB_Name = "modEngine"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit


' ==============================================================================
' Module: modEngine
' Description: clsTable과 내부 HashIndex/RowList를 활용한 고속 쿼리 연산 엔진
' Update: Semi-Join (존재 확인) 지원 추가
' ==============================================================================

' 조인 타입 정의
Public Enum joinType
    jtInner = 1      ' 교집합
    jtLeft = 2       ' 왼쪽 전체 + 교집합 (Left Outer)
    jtRight = 3      ' 교집합 + 오른쪽 전체 (Right Outer)
    jtFull = 4       ' 합집합 (Full Outer)
    jtLeftAnti = 5   ' 왼쪽 차집합 (Left Only) - B에 없는 A
    jtRightAnti = 6  ' 오른쪽 차집합 (Right Only) - A에 없는 B
    jtLeftSemi = 7   ' [NEW] 왼쪽 존재 확인 (Left Semi) - B에 있는 A
    jtRightSemi = 8  ' [NEW] 오른쪽 존재 확인 (Right Semi) - A에 있는 B
End Enum

' ==============================================================================
' Function: JoinTable (Universal Version + Header Safety)
' 설명: Inner, Outer, Anti, Semi Join을 모두 지원
' 실행: 탐침 1패스로 결과 행을 (왼쪽 rowId, 오른쪽 rowId) 쌍 목록에 기록한 뒤 결과 배열을
'       1회 할당하고 양쪽 GatherRows로 수집한다 (원본 전체 사본 없음).
' 행 순서 계약: 공통 키(오른쪽 커서 순, 키 내 왼쪽 x 오른쪽 중첩) -> 왼쪽 전용(왼쪽 커서 순)
'       -> 오른쪽 전용(오른쪽 커서 순). semi/anti는 해당 쪽만.
' ==============================================================================
Public Function JoinTable(ByVal tLeft As clsTable, _
                          ByVal tRight As clsTable, _
                          ByVal sLeftKey As String, _
                          ByVal sRightKey As String, _
                          Optional ByVal jType As joinType = jtInner, _
                          Optional ByVal sRightPrefix As String = "") As clsTable

    Dim vResultData As Variant
    Dim lRows As Long, lCols As Long
    Dim lLeftCols As Long, lRightCols As Long

    Dim i As Long
    Dim vKey As Variant

    Dim idxLeft As DbHashIndex, idxRight As DbHashIndex
    Dim cur As DbIndexCursor
    Dim bufLeft() As Long, bufRight() As Long ' span 재사용 버퍼 (자기조인 대비 좌우 분리)
    Dim cntLeft As Long, cntRight As Long
    Dim a As Long, b As Long
    Dim bHit As Boolean

    ' 행 쌍 목록 (0기반 증가형): 공통 쌍 -> 왼쪽 전용 -> 오른쪽 전용 순으로 출력
    Dim pairL() As Long, pairR() As Long, nPairs As Long, capPairL As Long, capPairR As Long
    Dim onlyL() As Long, nOnlyL As Long, capOnlyL As Long
    Dim onlyR() As Long, nOnlyR As Long, capOnlyR As Long
    Dim leftMatched() As Boolean ' 공통 패스에서 일치한 왼쪽 행 (왼쪽 전용 판정 - 재탐침 없음)
    Dim bCommon As Boolean, bLeftOnly As Boolean, bRightOnly As Boolean
    Dim bOutLeft As Boolean, bOutRight As Boolean
    Dim colsLeft() As Long, colsRight() As Long
    Dim vHdrLeft As Variant, vHdrRight As Variant

    ' 헤더 처리를 위한 변수
    Dim colHeaders As Collection
    Dim sHeader As String, sNewHeader As String, sTempName As String
    Dim nSuffix As Long

    ' v2.5.1 최적화 변수
    Dim lLeftKeyCol As Long, lRightKeyCol As Long

    ' 1. [Validation & Hoisting] 컬럼 인덱스 1회 확보
    ' 키 컬럼 오타는 문장 수준 오류 - fail-fast (audit 추보 F25). 이전 MsgBox는
    ' 헤드리스에서 모달 행 + 호출자에 Nothing 반환으로 조용한 91 전이를 일으켰다.
    If Not tLeft.HasColumn(sLeftKey) Then
        Err.Raise 9999, "modEngine.JoinTable", "Column not found (left join key): " & sLeftKey
    End If
    If Not tRight.HasColumn(sRightKey) Then
        Err.Raise 9999, "modEngine.JoinTable", "Column not found (right join key): " & sRightKey
    End If
    lLeftKeyCol = tLeft.ColumnIndex(sLeftKey)
    lRightKeyCol = tRight.ColumnIndex(sRightKey)

    ' 2. [Hash Index] 내부 실행용 인덱스 확보
    Set idxLeft = tLeft.GetHashIndexByIndex(lLeftKeyCol)
    Set idxRight = tRight.GetHashIndexByIndex(lRightKeyCol)

    bCommon = (jType = jtInner Or jType = jtLeft Or jType = jtRight Or jType = jtFull Or _
               jType = jtLeftSemi Or jType = jtRightSemi)
    bLeftOnly = (jType = jtLeft Or jType = jtFull Or jType = jtLeftAnti)
    bRightOnly = (jType = jtRight Or jType = jtFull Or jType = jtRightAnti)
    bOutLeft = (jType <> jtRightSemi And jType <> jtRightAnti)
    bOutRight = (jType <> jtLeftSemi And jType <> jtLeftAnti)

    ' 3. [Probe] 오른쪽 커서 1회전: 공통 키 행 쌍 + 오른쪽 전용 행
    If bCommon Or bRightOnly Then
        If bCommon And bLeftOnly Then ReDim leftMatched(0 To tLeft.RowCount)
        Set cur = idxRight.Cursor
        Do While cur.MoveNext
            vKey = cur.Key
            If bCommon And jType <> jtRightSemi Then
                bHit = idxLeft.TryGetRowIdSpan(vKey, bufLeft, cntLeft)
            Else
                bHit = idxLeft.TryGetRowCount(vKey, cntLeft)
            End If

            If bHit Then
                If jType = jtLeftSemi Then
                    If nPairs + cntLeft > capPairL Then GrowIds pairL, capPairL, nPairs + cntLeft
                    For a = 1 To cntLeft
                        pairL(nPairs) = bufLeft(a)
                        nPairs = nPairs + 1
                    Next a
                ElseIf jType = jtRightSemi Then
                    cur.CopyRows bufRight, cntRight
                    If nPairs + cntRight > capPairR Then GrowIds pairR, capPairR, nPairs + cntRight
                    For b = 1 To cntRight
                        pairR(nPairs) = bufRight(b)
                        nPairs = nPairs + 1
                    Next b
                ElseIf bCommon Then
                    cur.CopyRows bufRight, cntRight
                    lRows = nPairs + cntLeft * cntRight
                    If lRows > capPairL Then GrowIds pairL, capPairL, lRows
                    If lRows > capPairR Then GrowIds pairR, capPairR, lRows
                    For a = 1 To cntLeft
                        If bLeftOnly Then leftMatched(bufLeft(a)) = True
                        For b = 1 To cntRight
                            pairL(nPairs) = bufLeft(a)
                            pairR(nPairs) = bufRight(b)
                            nPairs = nPairs + 1
                        Next b
                    Next a
                End If
            ElseIf bRightOnly Then
                cur.CopyRows bufRight, cntRight
                If nOnlyR + cntRight > capOnlyR Then GrowIds onlyR, capOnlyR, nOnlyR + cntRight
                For b = 1 To cntRight
                    onlyR(nOnlyR) = bufRight(b)
                    nOnlyR = nOnlyR + 1
                Next b
            End If
        Loop
    End If

    ' 4. [Left Only] 왼쪽 커서 순. 공통 패스를 돌았으면 일치 플래그로 판정(키 내 행은 일치 여부 공유),
    '    left anti는 공통 패스가 없으므로 오른쪽 인덱스 탐침
    If bLeftOnly Then
        Set cur = idxLeft.Cursor
        Do While cur.MoveNext
            If bCommon Then
                cur.CopyRows bufLeft, cntLeft
                bHit = leftMatched(bufLeft(1))
            Else
                bHit = idxRight.TryGetRowCount(cur.Key, cntRight)
                If Not bHit Then cur.CopyRows bufLeft, cntLeft
            End If
            If Not bHit Then
                If nOnlyL + cntLeft > capOnlyL Then GrowIds onlyL, capOnlyL, nOnlyL + cntLeft
                For a = 1 To cntLeft
                    onlyL(nOnlyL) = bufLeft(a)
                    nOnlyL = nOnlyL + 1
                Next a
            End If
        Loop
    End If

    ' 5. [Allocation & Header Processing] 결과 배열 1회 할당 및 헤더 충돌 방지
    lLeftCols = tLeft.ColCount
    lRightCols = tRight.ColCount
    lRows = nPairs + nOnlyL + nOnlyR

    If Not bOutRight Then
        lCols = lLeftCols
    ElseIf Not bOutLeft Then
        lCols = lRightCols
    Else
        lCols = lLeftCols + lRightCols
    End If

    ReDim vResultData(1 To lRows + 1, 1 To lCols)
    Set colHeaders = New Collection

    ' --- (5-1) 왼쪽 헤더 등록 ---
    If bOutLeft Then
        vHdrLeft = tLeft.GetRecord(1)
        For i = 1 To lLeftCols
            sHeader = vHdrLeft(i)
            vResultData(1, i) = sHeader
            On Error Resume Next
            colHeaders.Add sHeader, UCase(sHeader)
            On Error GoTo 0
        Next i
    End If

    ' --- (5-2) 오른쪽 헤더 등록 ---
    If bOutRight Then
        Dim iStartCol As Long
        iStartCol = IIf(bOutLeft, lLeftCols, 0)

        vHdrRight = tRight.GetRecord(1)
        For i = 1 To lRightCols
            sHeader = vHdrRight(i)
            sNewHeader = sRightPrefix & sHeader

            If IsHeaderExists(colHeaders, sNewHeader) Then
                nSuffix = 1
                Do
                    sTempName = sNewHeader & "_R" & IIf(nSuffix = 1, "", CStr(nSuffix))
                    If Not IsHeaderExists(colHeaders, sTempName) Then
                        sNewHeader = sTempName
                        Exit Do
                    End If
                    nSuffix = nSuffix + 1
                Loop
            End If

            vResultData(1, iStartCol + i) = sNewHeader
            On Error Resume Next
            colHeaders.Add sNewHeader, UCase(sNewHeader)
            On Error GoTo 0
        Next i
    End If

    ' 6. [Gather] 행 쌍 목록대로 양쪽 원본에서 직접 수집 (빈 쪽 셀은 Empty 유지)
    If bOutLeft Then
        colsLeft = ColumnSeq(lLeftCols)
        If bOutRight Then
            tLeft.GatherRows pairL, nPairs, colsLeft, vResultData, 2, 1
            tLeft.GatherRows onlyL, nOnlyL, colsLeft, vResultData, 2 + nPairs, 1
        ElseIf jType = jtLeftSemi Then
            tLeft.GatherRows pairL, nPairs, colsLeft, vResultData, 2, 1
        Else
            tLeft.GatherRows onlyL, nOnlyL, colsLeft, vResultData, 2, 1
        End If
    End If
    If bOutRight Then
        colsRight = ColumnSeq(lRightCols)
        If bOutLeft Then
            tRight.GatherRows pairR, nPairs, colsRight, vResultData, 2, lLeftCols + 1
            tRight.GatherRows onlyR, nOnlyR, colsRight, vResultData, 2 + nPairs + nOnlyL, lLeftCols + 1
        ElseIf jType = jtRightSemi Then
            tRight.GatherRows pairR, nPairs, colsRight, vResultData, 2, 1
        Else
            tRight.GatherRows onlyR, nOnlyR, colsRight, vResultData, 2, 1
        End If
    End If

    ' 7. 결과 반환 (원본 셀은 이미 트리밍됨 - 입양 경로로 재트리밍 패스 생략)
    Dim tResult As New clsTable
    tResult.AdoptArray vResultData
    Set JoinTable = tResult

End Function

' [Internal] 0기반 rowId 버퍼를 need 이상으로 배가 확장 (cap = 현재 용량, 갱신됨)
Private Sub GrowIds(ByRef buf() As Long, ByRef cap As Long, ByVal need As Long)
    If need <= cap Then Exit Sub
    If cap < 64 Then cap = 64
    Do While cap < need
        cap = cap * 2
    Loop
    ReDim Preserve buf(0 To cap - 1)
End Sub

' [Internal] GatherRows용 열 번호 1..n (n < 1이면 미할당 - GatherRows가 n=0 행에서 읽지 않음)
Private Function ColumnSeq(ByVal n As Long) As Long()
    Dim cols() As Long, j As Long
    If n >= 1 Then
        ReDim cols(1 To n)
        For j = 1 To n: cols(j) = j: Next j
    End If
    ColumnSeq = cols
End Function

' ==============================================================================
' Helper: IsHeaderExists (기존 동일)
' ==============================================================================
Private Function IsHeaderExists(col As Collection, sKey As String) As Boolean
    On Error Resume Next
    IsHeaderExists = False
    Dim v As Variant
    v = col(UCase(sKey))
    If Err.Number = 0 Then IsHeaderExists = True
    On Error GoTo 0
End Function

' ==============================================================================
' [v2.5] JIT Optimization Utilities
' ==============================================================================

' Function: FastNorm (Search Criteria Sanitizer)
' 목적: 사용자 입력 조건(vValue)을 루프 진입 전 단 1회 정규화하여 '비교 준비 상태'로 만듦.
' 특징: UCase$를 제거하여 루프 내 메모리 할당이 없는 StrComp(..., vbTextCompare)와 함께 사용
' ==============================================================================
Public Function FastNorm(ByVal vValue As Variant) As String
    Select Case VarType(vValue)
        Case vbString
            FastNorm = Trim$(vValue)
        Case vbNull, vbEmpty, vbError
            FastNorm = ""
        Case Else ' Numbers, Dates, etc.
            FastNorm = CStr(vValue)
    End Select
End Function
