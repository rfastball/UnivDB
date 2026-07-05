Attribute VB_Name = "modEngine"
Option Explicit


' ==============================================================================
' Module: modEngine
' Description: clsTable과 clsSet을 활용한 고속 쿼리 연산 엔진 (v2.2)
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
                          
    Dim setLeft As clsSet, setRight As clsSet
    Dim vSets As Variant
    Dim setCommon As clsSet, setLeftOnly As clsSet, setRightOnly As clsSet
    
    Dim vResultData As Variant
    Dim lRows As Long, lCols As Long
    Dim lLeftCols As Long, lRightCols As Long
    
    Dim i As Long, j As Long
    Dim lRowIdx As Long
    Dim vKey As Variant
    
    Dim colLeftRows As Collection, colRightRows As Collection
    Dim vItemLeft As Variant, vItemRight As Variant
    Dim vLeftData As Variant, vRightData As Variant
    
    ' 헤더 처리를 위한 변수
    Dim colHeaders As Collection
    Dim sHeader As String, sNewHeader As String, sTempName As String
    Dim nSuffix As Long
    
    ' v2.5.1 최적화 변수
    Dim lLeftKeyCol As Long, lRightKeyCol As Long
    
    ' 1. [Validation & Hoisting] 컬럼 인덱스 1회 확보
    On Error Resume Next
    lLeftKeyCol = tLeft.Headers(UCase(sLeftKey))
    If Err.Number <> 0 Then MsgBox "Left Table Join Key 없음: " & sLeftKey, vbCritical: Exit Function
    
    lRightKeyCol = tRight.Headers(UCase(sRightKey))
    If Err.Number <> 0 Then MsgBox "Right Table Join Key 없음: " & sRightKey, vbCritical: Exit Function
    On Error GoTo 0
    
    ' 2. [Set Operation] 키 분류 (Single-Pass)
    ' 숫자 인덱스를 사용하여 조회 (v2.5.1)
    Set setLeft = New clsSet: setLeft.LoadFromArray tLeft.GetUniqueKeysByIndex(lLeftKeyCol)
    Set setRight = New clsSet: setRight.LoadFromArray tRight.GetUniqueKeysByIndex(lRightKeyCol)
    
    ' srAllSplit: 0=Intersect, 1=LeftOnly, 2=RightOnly
    vSets = setLeft.Compare(setRight, srAllSplit)
    
    Set setCommon = vSets(0)
    Set setLeftOnly = vSets(1)
    Set setRightOnly = vSets(2)
    
    ' 3. [Size Estimation] 조인 타입별 행 개수 계산
    lRows = 0
    
    ' (A) Common Keys (Inner, Outer, Semi)
    If (jType = jtInner Or jType = jtLeft Or jType = jtRight Or jType = jtFull Or _
        jType = jtLeftSemi Or jType = jtRightSemi) Then
        
        If setCommon.Count > 0 Then
            For Each vKey In setCommon.KeysArray
                If jType = jtLeftSemi Then
                    lRows = lRows + tLeft.SelectRowsByIndex(lLeftKeyCol, vKey).Count
                ElseIf jType = jtRightSemi Then
                    lRows = lRows + tRight.SelectRowsByIndex(lRightKeyCol, vKey).Count
                Else
                    lRows = lRows + (tLeft.SelectRowsByIndex(lLeftKeyCol, vKey).Count * _
                                     tRight.SelectRowsByIndex(lRightKeyCol, vKey).Count)
                End If
            Next vKey
        End If
    End If
    
    ' (B) Left Only Keys
    If (jType = jtLeft Or jType = jtFull Or jType = jtLeftAnti) Then
        If setLeftOnly.Count > 0 Then
            For Each vKey In setLeftOnly.KeysArray
                lRows = lRows + tLeft.SelectRowsByIndex(lLeftKeyCol, vKey).Count
            Next vKey
        End If
    End If
    
    ' (C) Right Only Keys
    If (jType = jtRight Or jType = jtFull Or jType = jtRightAnti) Then
        If setRightOnly.Count > 0 Then
            For Each vKey In setRightOnly.KeysArray
                lRows = lRows + tRight.SelectRowsByIndex(lRightKeyCol, vKey).Count
            Next vKey
        End If
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
    
    ' --- (4-1) 왼쪽 헤더 등록 ---
    If jType <> jtRightAnti And jType <> jtRightSemi Then
        For i = 1 To lLeftCols
            sHeader = tLeft.RawArray(1, i)
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
            sHeader = tRight.RawArray(1, i)
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
    vLeftData = tLeft.RawArray
    vRightData = tRight.RawArray
    lRowIdx = 2
    
    ' (Part 1) Common Keys
    If (jType = jtInner Or jType = jtLeft Or jType = jtRight Or jType = jtFull Or _
        jType = jtLeftSemi Or jType = jtRightSemi) And setCommon.Count > 0 Then
        
        For Each vKey In setCommon.KeysArray
            Set colLeftRows = tLeft.SelectRowsByIndex(lLeftKeyCol, vKey)
            Set colRightRows = tRight.SelectRowsByIndex(lRightKeyCol, vKey)
            
            If jType = jtLeftSemi Then
                For Each vItemLeft In colLeftRows
                    For j = 1 To lLeftCols: vResultData(lRowIdx, j) = vLeftData(vItemLeft, j): Next j
                    lRowIdx = lRowIdx + 1
                Next vItemLeft
                
            ElseIf jType = jtRightSemi Then
                For Each vItemRight In colRightRows
                    For j = 1 To lRightCols: vResultData(lRowIdx, j) = vRightData(vItemRight, j): Next j
                    lRowIdx = lRowIdx + 1
                Next vItemRight
                
            Else
                For Each vItemLeft In colLeftRows
                    For Each vItemRight In colRightRows
                        For j = 1 To lLeftCols: vResultData(lRowIdx, j) = vLeftData(vItemLeft, j): Next j
                        For j = 1 To lRightCols: vResultData(lRowIdx, lLeftCols + j) = vRightData(vItemRight, j): Next j
                        lRowIdx = lRowIdx + 1
                    Next vItemRight
                Next vItemLeft
            End If
        Next vKey
    End If
    
    ' (Part 2) Left Only Processing
    If (jType = jtLeft Or jType = jtFull Or jType = jtLeftAnti) And setLeftOnly.Count > 0 Then
        For Each vKey In setLeftOnly.KeysArray
            Set colLeftRows = tLeft.SelectRowsByIndex(lLeftKeyCol, vKey)
            For Each vItemLeft In colLeftRows
                For j = 1 To lLeftCols: vResultData(lRowIdx, j) = vLeftData(vItemLeft, j): Next j
                
                If jType <> jtLeftAnti Then
                    For j = 1 To lRightCols: vResultData(lRowIdx, lLeftCols + j) = Empty: Next j
                End If
                lRowIdx = lRowIdx + 1
            Next vItemLeft
        Next vKey
    End If
    
    ' (Part 3) Right Only Processing
    If (jType = jtRight Or jType = jtFull Or jType = jtRightAnti) And setRightOnly.Count > 0 Then
        For Each vKey In setRightOnly.KeysArray
            Set colRightRows = tRight.SelectRowsByIndex(lRightKeyCol, vKey)
            For Each vItemRight In colRightRows
                If jType <> jtRightAnti Then
                    For j = 1 To lLeftCols: vResultData(lRowIdx, j) = Empty: Next j
                    For j = 1 To lRightCols: vResultData(lRowIdx, lLeftCols + j) = vRightData(vItemRight, j): Next j
                Else
                    For j = 1 To lRightCols: vResultData(lRowIdx, j) = vRightData(vItemRight, j): Next j
                End If
                lRowIdx = lRowIdx + 1
            Next vItemRight
        Next vKey
    End If
    
    ' 6. 결과 반환
    Dim tResult As New clsTable
    tResult.LoadFromArray vResultData
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
