VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' Class: clsQuery
' 설명: clsTable에 대한 조건 검색을 최적화된 순서(Eq -> Range)로 수행
' 전략: 인덱스를 사용하는 동등조건(Where)을 먼저 수행하여 스캔 범위를 줄임
' ==============================================================================

Private m_SourceTable As clsTable
Private m_setCandidates As clsSet   ' 현재 살아남은 후보 행 번호(RowID) 집합
Private m_bIsFirstCondition As Boolean
Private m_dicUpdates As Object       ' 업데이트할 컬럼과 값 저장 (Set 절)

' [Projection & Sort & Limit]
Private m_sTargetColumns As String   ' 출력할 컬럼 목록 (쉼표 구분)
Private m_sSortColumn As String      ' 정렬할 컬럼명
Private m_bSortDesc As Boolean       ' 내림차순 여부
Private m_lLimit As Long             ' 반환할 최대 행 개수 (0이면 제한 없음)

Private Sub Class_Initialize()
    ' 기본 상태 설정
    m_bIsFirstCondition = True
    Set m_dicUpdates = CreateObject("Scripting.Dictionary")
    
    ' 기능별 기본값 초기화
    m_sTargetColumns = "*"           ' 기본: 전체 컬럼
    m_sSortColumn = ""               ' 기본: 정렬 없음
    m_bSortDesc = False              ' 기본: 오름차순
    m_lLimit = 0                     ' 기본: 개수 제한 없음
End Sub

Private Sub Class_Terminate()
    Set m_SourceTable = Nothing
    Set m_setCandidates = Nothing
    Set m_dicUpdates = Nothing
End Sub

' ------------------------------------------------------------------------------
' [Step 0] 쿼리 대상 테이블 설정
' ------------------------------------------------------------------------------
Public Function From(ByVal tTable As clsTable) As clsQuery
    Set m_SourceTable = tTable
    Set m_setCandidates = Nothing
    m_bIsFirstCondition = True
    m_dicUpdates.RemoveAll
    Set From = Me
End Function

' ------------------------------------------------------------------------------
' [Step 1] 동등 조건 (Equality) - 인덱스 활용 (Fast)
' 설명: 가장 먼저 호출하여 후보군을 대폭 줄이는 것을 권장
' ------------------------------------------------------------------------------
Public Function Where(ByVal sColName As String, ByVal vValue As Variant) As clsQuery
    Dim colRows As Collection
    Dim setStep As clsSet
    Dim vRow As Variant
    
    ' 1. 인덱스에서 행 번호 추출 (O(1))
    Set colRows = m_SourceTable.SelectRows(sColName, vValue)
    
    ' 2. Collection -> clsSet 변환
    Set setStep = New clsSet
    For Each vRow In colRows
        setStep.Add vRow
    Next vRow
    
    ' 3. 집합 연산 (교집합)
    If m_bIsFirstCondition Then
        Set m_setCandidates = setStep
        m_bIsFirstCondition = False
    Else
        ' 이미 후보군이 있다면 교집합(AND) 수행
        Set m_setCandidates = m_setCandidates.Intersection(setStep)
    End If
    
    Set Where = Me
End Function

' ==============================================================================
' [Scanning Filters] - 동작 확인된 AndFilter의 유전자를 그대로 이식
' ==============================================================================

' 1. 기본 필터 (v2.5 JIT Architecture)
' 설명: 연산자(sOp)에 따라 수치/문자 경로를 분리하고, 모든 전처리를 루프 외부로 격리 (Hoisting).
'       루프 내부의 함수 호출 및 메모리 할당을 제거하여 성능을 극대화.
Public Function Filter(ByVal sColName As String, ByVal sOp As String, ByVal vValue As Variant) As clsQuery
    Dim setSurvivors As New clsSet
    Dim vRowID As Variant
    Dim i As Long
    
    If m_bIsFirstCondition Then InitializeAllRows
    If m_setCandidates.Count = 0 Then Set Filter = Me: Exit Function
    
    ' 1. [Hoisting] 루프 외부에서 모든 전처리 완료
    Dim lColIdx As Long
    Dim vSource As Variant
    Dim vCandidates As Variant
    
    ' 1-1. 컬럼 인덱스 1회 조회
    On Error Resume Next
    lColIdx = m_SourceTable.Headers(UCase(sColName))
    If Err.Number <> 0 Then Err.Raise 9999, "clsQuery.Filter", "Column not found: " & sColName
    On Error GoTo 0
    
    ' 1-2. 소스 배열 로컬 변수에 할당 (클래스 인터페이스 호출 오버헤드 제거)
    vSource = m_SourceTable.RawArray
    vCandidates = m_setCandidates.KeysArray
    
    ' 2. [Dual Path] 연산자 종류에 따라 실행 경로 분기
    Select Case UCase(Trim(sOp))
        Case "=", "<>", ">", "<", ">=", "<="
            ' ==========================================================
            ' [Path A] 수치 연산 경로 (Numeric Path)
            ' ==========================================================
            Dim dTarget As Double
            Dim vCell As Variant
            
            ' 비교값 사전 변환
            If Not IsNumeric(vValue) Then ' 비교값이 숫자가 아니면 필터링 결과 0
                Set m_setCandidates = New clsSet
                Set Filter = Me
                Exit Function
            End If
            dTarget = CDbl(vValue)
            
            ' 최적화된 루프
            For i = LBound(vCandidates) To UBound(vCandidates)
                vCell = vSource(vCandidates(i), lColIdx)
                If IsNumeric(vCell) Then
                    Select Case UCase(Trim(sOp))
                        Case "=": If CDbl(vCell) = dTarget Then setSurvivors.Add vCandidates(i)
                        Case "<>": If CDbl(vCell) <> dTarget Then setSurvivors.Add vCandidates(i)
                        Case ">": If CDbl(vCell) > dTarget Then setSurvivors.Add vCandidates(i)
                        Case "<": If CDbl(vCell) < dTarget Then setSurvivors.Add vCandidates(i)
                        Case ">=": If CDbl(vCell) >= dTarget Then setSurvivors.Add vCandidates(i)
                        Case "<=": If CDbl(vCell) <= dTarget Then setSurvivors.Add vCandidates(i)
                    End Select
                End If
            Next i
            
        Case "IS", "IS NOT", "LIKE", "NOT LIKE"
            ' ==========================================================
            ' [Path B] 문자열 연산 경로 (String Path)
            ' ==========================================================
            Dim sTarget As String
            sTarget = FastNorm(vValue) ' FastNorm으로 비교값 사전 정규화
            
            Select Case UCase(Trim(sOp))
                Case "IS" ' StrComp 기반의 대소문자 무시 비교 (=)
                    For i = LBound(vCandidates) To UBound(vCandidates)
                        If StrComp(CStr(vSource(vCandidates(i), lColIdx)), sTarget, vbTextCompare) = 0 Then
                            setSurvivors.Add vCandidates(i)
                        End If
                    Next i
                Case "IS NOT" ' StrComp 기반의 대소문자 무시 비교 (<>)
                    For i = LBound(vCandidates) To UBound(vCandidates)
                        If StrComp(CStr(vSource(vCandidates(i), lColIdx)), sTarget, vbTextCompare) <> 0 Then
                            setSurvivors.Add vCandidates(i)
                        End If
                    Next i
                Case "LIKE" ' VBA Like 연산자
                    sTarget = UCase(sTarget) ' Like는 vbTextCompare를 지원 안하므로 UCase 필요
                    For i = LBound(vCandidates) To UBound(vCandidates)
                        If UCase(CStr(vSource(vCandidates(i), lColIdx))) Like sTarget Then
                            setSurvivors.Add vCandidates(i)
                        End If
                    Next i
                Case "NOT LIKE"
                     sTarget = UCase(sTarget)
                     For i = LBound(vCandidates) To UBound(vCandidates)
                        If Not (UCase(CStr(vSource(vCandidates(i), lColIdx))) Like sTarget) Then
                            setSurvivors.Add vCandidates(i)
                        End If
                    Next i
            End Select
    End Select
    
    Set m_setCandidates = setSurvivors
    m_bIsFirstCondition = False
    Set Filter = Me
End Function

' ==============================================================================
' [Step 2.5] Projection 설정 (SQL의 SELECT 절)
' 설명: 결과로 출력할 컬럼과 그 순서를 지정합니다. (기본값: "*")
' 사용: .SelectColumns("Name, Dept, Age")
' ------------------------------------------------------------------------------
Public Function SelectColumns(ByVal sCols As String) As clsQuery
    If Len(Trim(sCols)) = 0 Then sCols = "*"
    m_sTargetColumns = sCols
    Set SelectColumns = Me
End Function

' ==============================================================================
' [Step 2.8] 정렬 설정 (SQL의 ORDER BY)
' 설명: SelectTable() 실행 시 반환될 결과 테이블을 정렬하도록 예약합니다.
' 활용: 기존 clsTable의 QuickSort 엔진을 재사용합니다.
' ==============================================================================
Public Function OrderBy(ByVal sColName As String, Optional ByVal bDesc As Boolean = False) As clsQuery
    m_sSortColumn = sColName
    m_bSortDesc = bDesc
    Set OrderBy = Me
End Function
' ==============================================================================
' [Step 2.9] 결과 제한 (SQL의 TOP / LIMIT)
' 설명: 반환할 최대 행 개수를 설정합니다.
' ==============================================================================
Public Function Top(ByVal lCount As Long) As clsQuery
    If lCount < 0 Then lCount = 0
    m_lLimit = lCount
    Set Top = Me
End Function

' ------------------------------------------------------------------------------
' [Step 2.6] Update 값 설정 (SQL SET)
' 설명: 업데이트할 컬럼과 값을 설정합니다. (체이닝 지원)
' ------------------------------------------------------------------------------
Public Function SetValue(ByVal sColName As String, ByVal vValue As Variant) As clsQuery
    If m_dicUpdates.exists(sColName) Then
        m_dicUpdates(sColName) = vValue
    Else
        m_dicUpdates.Add sColName, vValue
    End If
    Set SetValue = Me
End Function

' ------------------------------------------------------------------------------
' [Step 4] Update 실행 (SQL UPDATE)
' 설명: 조건에 맞는 행들의 컬럼 값을 일괄 변경합니다.
' 반환: 변경된 행의 개수
' ------------------------------------------------------------------------------
Public Function Update() As Long
    Dim vRowID As Variant, vCol As Variant
    Dim lCount As Long
    
    If m_setCandidates Is Nothing Then InitializeAllRows
    If m_dicUpdates.Count = 0 Or m_setCandidates.Count = 0 Then Update = 0: Exit Function
    
    For Each vRowID In m_setCandidates.KeysArray
        For Each vCol In m_dicUpdates.Keys
            m_SourceTable.Value(CLng(vRowID), CStr(vCol)) = m_dicUpdates(vCol)
        Next vCol
        lCount = lCount + 1
    Next vRowID
    
    Update = lCount
End Function

' ------------------------------------------------------------------------------
' [Step 4.1] Delete 실행 (SQL DELETE)
' 설명: 조건에 맞는 행을 테이블에서 삭제합니다.
' 반환: 삭제된 행의 개수
' ------------------------------------------------------------------------------
Public Function Delete() As Long
    Dim vSource As Variant
    Dim vNewData As Variant
    Dim vRowIDs As Variant
    Dim dicToDelete As Object
    Dim i As Long, j As Long, R As Long
    Dim lTotalRows As Long, lTotalCols As Long
    Dim lDeleteCount As Long
    Dim lKeepCount As Long
    
    If m_setCandidates Is Nothing Then InitializeAllRows
    lDeleteCount = m_setCandidates.Count
    If lDeleteCount = 0 Then Delete = 0: Exit Function
    
    ' 1. 삭제할 RowID 마킹
    Set dicToDelete = CreateObject("Scripting.Dictionary")
    vRowIDs = m_setCandidates.KeysArray
    If Not IsEmpty(vRowIDs) Then
        For i = LBound(vRowIDs) To UBound(vRowIDs)
            dicToDelete(vRowIDs(i)) = Empty
        Next i
    End If
    
    ' 2. 데이터 재구성 (삭제 대상 제외)
    vSource = m_SourceTable.RawArray
    lTotalRows = UBound(vSource, 1)
    lTotalCols = UBound(vSource, 2)
    lKeepCount = lTotalRows - lDeleteCount
    
    ReDim vNewData(1 To lKeepCount, 1 To lTotalCols)
    
    R = 1
    For j = 1 To lTotalCols: vNewData(1, j) = vSource(1, j): Next j ' 헤더
    
    For i = 2 To lTotalRows
        If Not dicToDelete.exists(i) Then
            R = R + 1
            For j = 1 To lTotalCols: vNewData(R, j) = vSource(i, j): Next j
        End If
    Next i
    
    ' 3. 테이블 갱신
    m_SourceTable.LoadFromArray vNewData, True
    Set m_setCandidates = Nothing: m_bIsFirstCondition = True
    
    Delete = lDeleteCount
End Function

' ------------------------------------------------------------------------------
' [Step 4.2] Insert 실행 (SQL INSERT)
' 설명: SetValue로 설정된 값으로 새로운 행을 추가합니다.
' 반환: 추가된 행의 개수 (성공 시 1)
' ------------------------------------------------------------------------------
Public Function Insert() As Long
    Dim vSource As Variant
    Dim vNewData As Variant
    Dim lTotalRows As Long, lTotalCols As Long
    Dim i As Long, j As Long
    Dim dicHeaders As Object
    Dim vKey As Variant
    Dim sCol As String
    Dim lColIdx As Long
    
    If m_dicUpdates.Count = 0 Then Insert = 0: Exit Function
    
    ' 1. 원본 데이터 및 헤더 확보
    vSource = m_SourceTable.RawArray
    lTotalRows = UBound(vSource, 1)
    lTotalCols = UBound(vSource, 2)
    Set dicHeaders = m_SourceTable.Headers
    
    ' 2. 배열 확장 (행 추가) 및 데이터 복사
    ReDim vNewData(1 To lTotalRows + 1, 1 To lTotalCols)
    For i = 1 To lTotalRows
        For j = 1 To lTotalCols: vNewData(i, j) = vSource(i, j): Next j
    Next i
    
    ' 3. 새 데이터 매핑
    For Each vKey In m_dicUpdates.Keys
        sCol = UCase(Trim(CStr(vKey)))
        If dicHeaders.exists(sCol) Then
            lColIdx = dicHeaders(sCol)
            vNewData(lTotalRows + 1, lColIdx) = m_dicUpdates(vKey)
        Else
            Err.Raise 9999, "clsQuery", "Insert Error: Column not found [" & sCol & "]"
        End If
    Next vKey
    
    ' 4. 테이블 갱신
    m_SourceTable.LoadFromArray vNewData, True
    m_dicUpdates.RemoveAll
    
    Insert = 1
End Function

' ------------------------------------------------------------------------------
' [Step 3] 최종 실행 및 결과 반환 (Projection + Top/Limit + OrderBy 통합)
' ------------------------------------------------------------------------------
Public Function SelectTable() As clsTable
    Dim tResult As New clsTable
    Dim vResultData As Variant
    Dim vRowIDs As Variant
    Dim vRawSource As Variant
    
    Dim arrColNames() As String
    Dim arrSourceIndices() As Long ' Target Index -> Source Index 매핑
    Dim lTargetCols As Long
    Dim lResultRows As Long
    
    Dim i As Long, j As Long, R As Long
    Dim sCol As String
    
    ' 1. [Candidate Check] 후보군 확인
    If m_setCandidates Is Nothing Then
        InitializeAllRows
    End If
    lResultRows = m_setCandidates.Count
    
    ' [Optimized Limit] 정렬이 없는 경우, 미리 개수를 제한하여 성능 최적화 (Early Exit)
    If Len(m_sSortColumn) = 0 And m_lLimit > 0 Then
        If lResultRows > m_lLimit Then lResultRows = m_lLimit
    End If
    
    ' 2. [Column Parsing] 프로젝션 대상 컬럼 분석
    Dim dicHeaders As Object
    Set dicHeaders = m_SourceTable.Headers
    
    If m_sTargetColumns = "*" Or m_sTargetColumns = "" Then
        ' [Case A] 전체 컬럼 선택
        lTargetCols = m_SourceTable.ColCount
        ReDim arrSourceIndices(1 To lTargetCols)
        For j = 1 To lTargetCols: arrSourceIndices(j) = j: Next j
    Else
        ' [Case B] 특정 컬럼 선택
        arrColNames = Split(m_sTargetColumns, ",")
        lTargetCols = UBound(arrColNames) + 1
        ReDim arrSourceIndices(1 To lTargetCols)
        
        For j = 0 To UBound(arrColNames)
            sCol = UCase(Trim(arrColNames(j)))
            If dicHeaders.exists(sCol) Then
                arrSourceIndices(j + 1) = dicHeaders(sCol)
            Else
                Err.Raise 9999, , "SelectColumns Error: Column not found [" & sCol & "]"
            End If
        Next j
    End If
    
    ' 3. [Allocation] 결과 배열 할당 (Limit이 적용된 lResultRows 사용)
    ReDim vResultData(1 To lResultRows + 1, 1 To lTargetCols)
    vRawSource = m_SourceTable.RawArray
    
    ' 4. [Header Copy] 헤더 작성
    If m_sTargetColumns = "*" Then
        For j = 1 To lTargetCols: vResultData(1, j) = vRawSource(1, j): Next j
    Else
        For j = 0 To UBound(arrColNames): vResultData(1, j + 1) = Trim(arrColNames(j)): Next j
    End If
    
    ' 5. [Data Fetch] 데이터 고속 복사
    If lResultRows > 0 Then
        vRowIDs = m_setCandidates.KeysArray
        
        ' [핵심] lResultRows만큼만 루프 수행 (정렬 없는 경우 Limit까지만 돎)
        ' vRowIDs는 0-based array이므로 0부터 시작
        For i = 0 To lResultRows - 1
            R = vRowIDs(i)
            For j = 1 To lTargetCols
                vResultData(i + 2, j) = vRawSource(R, arrSourceIndices(j))
            Next j
        Next i
    End If
    
    ' 결과 로드
    tResult.LoadFromArray vResultData, True
    
    ' 6. [Post-Process] 정렬 및 정렬 후 Limit 처리
    If Len(m_sSortColumn) > 0 Then
        ' (1) 정렬 수행 (전체 데이터 대상)
        If tResult.Headers.exists(UCase(m_sSortColumn)) Then
            tResult.OrderBy m_sSortColumn, m_bSortDesc
        End If
        
        ' (2) 정렬 후 잘라내기 (Slicing) - 정렬이 있을 때만 수행되는 후처리
        If m_lLimit > 0 And tResult.RowCount > m_lLimit Then
            Dim vSliced As Variant
            Dim vSorted As Variant
            
            vSorted = tResult.RawArray
            ReDim vSliced(1 To m_lLimit + 1, 1 To lTargetCols)
            
            ' 헤더 + 상위 N개 복사
            For i = 1 To m_lLimit + 1
                For j = 1 To lTargetCols
                    vSliced(i, j) = vSorted(i, j)
                Next j
            Next i
            
            ' 잘라낸 데이터로 다시 로드
            tResult.LoadFromArray vSliced, True
        End If
    End If
    
    Set SelectTable = tResult
End Function

' ==============================================================================
' [Internal] 헬퍼 함수
' ==============================================================================

' 전체 행을 후보군으로 초기화 (Where 없이 Filter 먼저 호출 시)
Private Sub InitializeAllRows()
    Dim i As Long
    Set m_setCandidates = New clsSet
    For i = 2 To m_SourceTable.RowCount ' 1은 헤더
        m_setCandidates.Add i
    Next i
    m_bIsFirstCondition = False
End Sub

