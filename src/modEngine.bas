Attribute VB_Name = "modEngine"
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
    
    Dim i As Long, j As Long
    Dim lRowIdx As Long
    Dim vKey As Variant
    
    Dim idxLeft As DbHashIndex, idxRight As DbHashIndex
    Dim cur As DbIndexCursor
    Dim bufLeft() As Long, bufRight() As Long ' span 재사용 버퍼 (자기조인 대비 좌우 분리)
    Dim cntLeft As Long, cntRight As Long
    Dim vItemLeft As Long, vItemRight As Long
    Dim a As Long, b As Long
    Dim vLeftData As Variant, vRightData As Variant
    
    ' 헤더 처리를 위한 변수
    Dim colHeaders As Collection
    Dim sHeader As String, sNewHeader As String, sTempName As String
    Dim nSuffix As Long
    
    ' v2.5.1 최적화 변수
    Dim lLeftKeyCol As Long, lRightKeyCol As Long
    
    ' 1. [Validation & Hoisting] 컬럼 인덱스 1회 확보
    ' 키 컬럼 오타는 문장 수준 오류 - fail-fast (audit 추보 F25). 이전 MsgBox는
    ' 헤드리스에서 모달 행 + 호출자에 Nothing 반환으로 조용한 91 전이를 일으켰다.
    If Not tLeft.Headers.exists(UCase(sLeftKey)) Then
        Err.Raise 9999, "modEngine.JoinTable", "Column not found (left join key): " & sLeftKey
    End If
    If Not tRight.Headers.exists(UCase(sRightKey)) Then
        Err.Raise 9999, "modEngine.JoinTable", "Column not found (right join key): " & sRightKey
    End If
    lLeftKeyCol = tLeft.Headers(UCase(sLeftKey))
    lRightKeyCol = tRight.Headers(UCase(sRightKey))
    
    ' 2. [Hash Index] 내부 실행용 인덱스 확보
    Set idxLeft = tLeft.GetHashIndexByIndex(lLeftKeyCol)
    Set idxRight = tRight.GetHashIndexByIndex(lRightKeyCol)
    
    ' 3. [Size Estimation] 조인 타입별 행 개수 계산
    lRows = 0
    
    ' (A) Common Keys (Inner, Outer, Semi)
    If (jType = jtInner Or jType = jtLeft Or jType = jtRight Or jType = jtFull Or _
        jType = jtLeftSemi Or jType = jtRightSemi) Then
        
        Set cur = idxRight.Cursor
        Do While cur.MoveNext
            If idxLeft.TryGetRowCount(cur.Key, cntLeft) Then
                If jType = jtLeftSemi Then
                    lRows = lRows + cntLeft
                ElseIf jType = jtRightSemi Then
                    lRows = lRows + cur.RowCount
                Else
                    lRows = lRows + (cntLeft * cur.RowCount)
                End If
            End If
        Loop
    End If

    ' (B) Left Only Keys
    If (jType = jtLeft Or jType = jtFull Or jType = jtLeftAnti) Then
        Set cur = idxLeft.Cursor
        Do While cur.MoveNext
            If Not idxRight.TryGetRowCount(cur.Key, cntRight) Then
                lRows = lRows + cur.RowCount
            End If
        Loop
    End If

    ' (C) Right Only Keys
    If (jType = jtRight Or jType = jtFull Or jType = jtRightAnti) Then
        Set cur = idxRight.Cursor
        Do While cur.MoveNext
            If Not idxLeft.TryGetRowCount(cur.Key, cntLeft) Then
                lRows = lRows + cur.RowCount
            End If
        Loop
    End If
    
    ' 4. [Allocation & Header Processing] 결과 배열 생성 및 헤더 충돌 방지
    lLeftCols = tLeft.ColCount
    lRightCols = tRight.ColCount
    
    If jType = jtLeftAnti Or jType = jtLeftSemi Then
        lCols = lLeftCols
    ElseIf jType = jtRightAnti Or jType = jtRightSemi Then
        lCols = lRightCols
    Else
        lCols = lLeftCols + lRightCols
    End If
    
    ReDim vResultData(1 To lRows + 1, 1 To lCols)
    Set colHeaders = New Collection

    ' 원본 배열 1회 물질화 (RawArray는 호출마다 전체 복사 - 루프 안 호출 금지)
    vLeftData = tLeft.RawArray
    vRightData = tRight.RawArray

    ' --- (4-1) 왼쪽 헤더 등록 ---
    If jType <> jtRightAnti And jType <> jtRightSemi Then
        For i = 1 To lLeftCols
            sHeader = vLeftData(1, i)
            vResultData(1, i) = sHeader
            On Error Resume Next
            colHeaders.Add sHeader, UCase(sHeader)
            On Error GoTo 0
        Next i
    End If
    
    ' --- (4-2) 오른쪽 헤더 등록 ---
    If jType <> jtLeftAnti And jType <> jtLeftSemi Then
        Dim iStartCol As Long
        iStartCol = IIf(jType = jtRightAnti Or jType = jtRightSemi, 0, lLeftCols)
        
        For i = 1 To lRightCols
            sHeader = vRightData(1, i)
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
    
    ' 5. [Data Merge] 데이터 채우기
    lRowIdx = 2
    
    ' (Part 1) Common Keys
    If (jType = jtInner Or jType = jtLeft Or jType = jtRight Or jType = jtFull Or _
        jType = jtLeftSemi Or jType = jtRightSemi) Then
        
        Set cur = idxRight.Cursor
        Do While cur.MoveNext
            vKey = cur.Key
            If Not idxLeft.TryGetRowIdSpan(vKey, bufLeft, cntLeft) Then GoTo NextCommonKey

            If jType = jtLeftSemi Then
                For a = 1 To cntLeft
                    vItemLeft = bufLeft(a)
                    For j = 1 To lLeftCols: vResultData(lRowIdx, j) = vLeftData(vItemLeft, j): Next j
                    lRowIdx = lRowIdx + 1
                Next a

            ElseIf jType = jtRightSemi Then
                cur.CopyRows bufRight, cntRight
                For b = 1 To cntRight
                    vItemRight = bufRight(b)
                    For j = 1 To lRightCols: vResultData(lRowIdx, j) = vRightData(vItemRight, j): Next j
                    lRowIdx = lRowIdx + 1
                Next b

            Else
                cur.CopyRows bufRight, cntRight
                For a = 1 To cntLeft
                    vItemLeft = bufLeft(a)
                    For b = 1 To cntRight
                        vItemRight = bufRight(b)
                        For j = 1 To lLeftCols: vResultData(lRowIdx, j) = vLeftData(vItemLeft, j): Next j
                        For j = 1 To lRightCols: vResultData(lRowIdx, lLeftCols + j) = vRightData(vItemRight, j): Next j
                        lRowIdx = lRowIdx + 1
                    Next b
                Next a
            End If
NextCommonKey:
        Loop
    End If
    
    ' (Part 2) Left Only Processing
    If (jType = jtLeft Or jType = jtFull Or jType = jtLeftAnti) Then
        Set cur = idxLeft.Cursor
        Do While cur.MoveNext
            vKey = cur.Key
            If idxRight.TryGetRowCount(vKey, cntRight) Then GoTo NextLeftOnlyKey
            cur.CopyRows bufLeft, cntLeft
            For a = 1 To cntLeft
                vItemLeft = bufLeft(a)
                For j = 1 To lLeftCols: vResultData(lRowIdx, j) = vLeftData(vItemLeft, j): Next j

                If jType <> jtLeftAnti Then
                    For j = 1 To lRightCols: vResultData(lRowIdx, lLeftCols + j) = Empty: Next j
                End If
                lRowIdx = lRowIdx + 1
            Next a
NextLeftOnlyKey:
        Loop
    End If
    
    ' (Part 3) Right Only Processing
    If (jType = jtRight Or jType = jtFull Or jType = jtRightAnti) Then
        Set cur = idxRight.Cursor
        Do While cur.MoveNext
            vKey = cur.Key
            If idxLeft.TryGetRowCount(vKey, cntLeft) Then GoTo NextRightOnlyKey
            cur.CopyRows bufRight, cntRight
            For b = 1 To cntRight
                vItemRight = bufRight(b)
                If jType <> jtRightAnti Then
                    For j = 1 To lLeftCols: vResultData(lRowIdx, j) = Empty: Next j
                    For j = 1 To lRightCols: vResultData(lRowIdx, lLeftCols + j) = vRightData(vItemRight, j): Next j
                Else
                    For j = 1 To lRightCols: vResultData(lRowIdx, j) = vRightData(vItemRight, j): Next j
                End If
                lRowIdx = lRowIdx + 1
            Next b
NextRightOnlyKey:
        Loop
    End If
    
    ' 6. 결과 반환 (원본 셀은 이미 트리밍됨 - 입양 경로로 재트리밍 패스 생략)
    Dim tResult As New clsTable
    tResult.AdoptArray vResultData
    Set JoinTable = tResult
    
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
