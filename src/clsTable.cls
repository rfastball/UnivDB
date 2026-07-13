VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ==============================================================================
' Class Name: clsTable
' Description: 엑셀 시트 데이터를 메모리화하여 관리하는 고속 데이터 컨테이너
' Features:
'   1. Lazy Indexing: 검색이 필요한 시점에 인덱스 생성 (메모리 절약)
'   2. Row-ID Mapping: 중복 데이터 허용 (Key -> Collection of RowIDs)
'   3. Direct Access: 2D 배열 기반의 초고속 데이터 접근
' ==============================================================================

' [저장소]
Private m_vData As Variant          ' 전체 데이터 (2D Array, 1-based or 0-based)
Private m_dicHeaders As Object      ' 컬럼명 -> 열 번호 매핑 (Key: String, Value: Long)
Private m_dicHashIndices As Object  ' 인덱스 캐시 (Key: ColName/Index, Value: DbHashIndex) - 유일한 인덱스 진실
Private m_lIndexBackend As DbIndexBackend ' 인덱스 백엔드 선택 (기본 dbBackendDict, 자동 선택 없음)

' [메타 데이터]
Private m_lRowCount As Long
Private m_lColCount As Long
Private m_lHeaderRowIndex As Long   ' 헤더가 위치한 행 번호 (기본 1)
Private m_lGeneration As Long       ' 구조 변경(행 수/순서/컬럼 배치)마다 증가하는 세대 번호

' [초기화]
Private Sub Class_Initialize()
    Set m_dicHeaders = CreateObject("Scripting.Dictionary")
    Set m_dicHashIndices = CreateObject("Scripting.Dictionary")
    m_lHeaderRowIndex = 1
End Sub

Private Sub Class_Terminate()
    Set m_dicHeaders = Nothing
    Set m_dicHashIndices = Nothing
    m_vData = Empty
End Sub

' ==============================================================================
' [1] 데이터 로드 (Load)
' ==============================================================================

' ==============================================================================
' [1-1] 시트 데이터 로드 (Smart Range Detection 적용)
' 사용법 1: t.LoadFromSheet Sheets("Data")        -> 시트 전체 데이터 자동 탐색 (추천)
' 사용법 2: t.LoadFromSheet Range("A1")           -> A1부터 끝까지 자동 탐색 (추천)
' 사용법 3: t.LoadFromSheet Range("A1:F500")      -> 지정된 범위만 강제 로드
' ==============================================================================
Public Sub LoadFromSheet(ByVal vSource As Variant)
    Dim rngTarget As Range
    
    ' 1. 입력 타입 확인 및 범위 확정
    If TypeOf vSource Is Worksheet Then
        ' [Case A] 시트 객체가 넘어온 경우 -> 데이터 있는 전체 범위 자동 탐색
        Set rngTarget = GetRealUsedRange(vSource)
        
    ElseIf TypeOf vSource Is Range Then
        ' [Case B] Range 객체가 넘어온 경우
        If vSource.Cells.Count = 1 Then
            ' [Case B-1] 단일 셀(예: A1)만 넘김 -> 여기서부터 끝까지 자동 탐색
            Set rngTarget = GetRealUsedRange(vSource.Worksheet, vSource)
        Else
            ' [Case B-2] 구체적 범위(예: A1:E100)를 넘김 -> CurrentRegion 안 쓰고 그대로 믿음
            Set rngTarget = vSource
        End If
    Else
        Err.Raise 9999, , "LoadFromSheet: Worksheet 또는 Range 객체만 지원합니다."
    End If
    
    ' 2. 데이터 유효성 검사
    If rngTarget Is Nothing Then
        m_lRowCount = 0
        Exit Sub
    End If
    
    ' 3. 메모리 적재 (배열로 변환)
    ' * Value2가 Value보다 속도가 빠름 (통화/날짜 변환 오버헤드 없음)
    Dim vData As Variant
    vData = rngTarget.Value2
    
    ' 4. 배열 로드 실행
    LoadFromArray vData
End Sub
' ==============================================================================
' [Best Practice] Find 메서드를 이용한 고속 범위 탐색
' ==============================================================================
Private Function GetRealUsedRange(ByVal ws As Worksheet, Optional ByVal rngStart As Range = Nothing) As Range
    Dim lLastRow As Long, lLastCol As Long
    Dim lStartRow As Long, lStartCol As Long
    Dim rngFindRow As Range, rngFindCol As Range
    
    ' 1. 시작점 설정 (기본 A1)
    If rngStart Is Nothing Then
        lStartRow = 1: lStartCol = 1
    Else
        lStartRow = rngStart.Row: lStartCol = rngStart.Column
    End If
    
    ' 2. 데이터가 있는 '진짜' 마지막 행 찾기 (뒤에서부터 검색)
    ' * xlFormulas 옵션: 숨겨진 수식이나 값 모두 감지
    Set rngFindRow = ws.Cells.Find(What:="*", _
                                   After:=ws.Cells(1, 1), _
                                   LookIn:=xlFormulas, _
                                   LookAt:=xlPart, _
                                   SearchOrder:=xlByRows, _
                                   SearchDirection:=xlPrevious)
                                   
    ' 3. 데이터가 있는 '진짜' 마지막 열 찾기
    Set rngFindCol = ws.Cells.Find(What:="*", _
                                   After:=ws.Cells(1, 1), _
                                   LookIn:=xlFormulas, _
                                   LookAt:=xlPart, _
                                   SearchOrder:=xlByColumns, _
                                   SearchDirection:=xlPrevious)
    
    ' 4. 데이터가 아예 없는 빈 시트일 경우 방어
    If rngFindRow Is Nothing Or rngFindCol Is Nothing Then
        ' 데이터 없음 -> 시작 셀 하나만 반환 (에러 방지)
        Set GetRealUsedRange = ws.Cells(lStartRow, lStartCol)
        Exit Function
    End If
    
    lLastRow = rngFindRow.Row
    lLastCol = rngFindCol.Column
    
    ' 시작점보다 데이터가 앞쪽에 있으면 보정
    If lLastRow < lStartRow Then lLastRow = lStartRow
    If lLastCol < lStartCol Then lLastCol = lStartCol
    
    ' 5. 사각형 범위 확정 (중간에 구멍이 뚫려도 직사각형으로 잡음)
    Set GetRealUsedRange = ws.Range(ws.Cells(lStartRow, lStartCol), ws.Cells(lLastRow, lLastCol))
End Function

' 1-2. 배열에서 직접 로드 (데이터 정규화 포함 버전)
Public Sub LoadFromArray(ByVal vArr As Variant, Optional ByVal bHasHeader As Boolean = True)
    Dim i As Long, j As Long
    Dim sColName As String
    
    Clear ' 기존 데이터 초기화
    
    m_vData = vArr
    m_lRowCount = UBound(m_vData, 1)
    m_lColCount = UBound(m_vData, 2)
    
    ' [전처리] 전체 데이터 영역 초고속 트리밍
    ' * 문자열 타입(vbString)인 경우에만 Trim$ 수행하여 속도 최적화
    For i = LBound(m_vData, 1) To UBound(m_vData, 1)
        For j = LBound(m_vData, 2) To UBound(m_vData, 2)
            If VarType(m_vData(i, j)) = vbString Then
                m_vData(i, j) = Trim$(m_vData(i, j))
            End If
        Next j
    Next i
    
    ' 헤더 매핑 생성
    If bHasHeader Then
        Dim lHeadRow As Long: lHeadRow = LBound(m_vData, 1)
        For j = LBound(m_vData, 2) To UBound(m_vData, 2)
            ' 헤더 자체도 다시 한번 Trim (이미 위에서 처리되었지만 명시적 유지)
            sColName = CStr(m_vData(lHeadRow, j))
            If Len(sColName) > 0 Then
                m_dicHeaders(UCase(sColName)) = j
            End If
        Next j
        m_lHeaderRowIndex = lHeadRow
    Else
        ' 헤더가 없는 경우
        For j = LBound(m_vData, 2) To UBound(m_vData, 2)
            m_dicHeaders("COLUMN" & j) = j
        Next j
        m_lHeaderRowIndex = LBound(m_vData, 1) - 1
    End If
End Sub

' 1-3. [Internal] 이미 정규화된 배열 입양 (내부 물질화 전용)
' LoadFromArray와 달리 데이터 전수 Trim 패스를 생략하고 ByRef로 받아 호출 시점 복사를 줄인다.
' 원본이 이미 트리밍된 테이블 셀/집계값인 내부 결과(조인, GroupBy, 쿼리 재구축)에서만 쓸 것.
' 사용자 입력이 섞이는 경로(RegisterTable, Insert, TableBuilder)는 LoadFromArray를 유지한다.
Friend Sub AdoptArray(ByRef vArr As Variant, Optional ByVal bHasHeader As Boolean = True)
    Dim j As Long
    Dim sColName As String

    Clear ' 기존 데이터 초기화 (세대 증가 포함 - 파생 구조 실효)

    m_vData = vArr
    m_lRowCount = UBound(m_vData, 1)
    m_lColCount = UBound(m_vData, 2)

    If bHasHeader Then
        ' 헤더 행만 Trim (헤더 이름 의미론은 LoadFromArray와 동일 유지)
        Dim lHeadRow As Long: lHeadRow = LBound(m_vData, 1)
        For j = LBound(m_vData, 2) To UBound(m_vData, 2)
            If VarType(m_vData(lHeadRow, j)) = vbString Then
                m_vData(lHeadRow, j) = Trim$(m_vData(lHeadRow, j))
            End If
            sColName = CStr(m_vData(lHeadRow, j))
            If Len(sColName) > 0 Then
                m_dicHeaders(UCase(sColName)) = j
            End If
        Next j
        m_lHeaderRowIndex = lHeadRow
    Else
        For j = LBound(m_vData, 2) To UBound(m_vData, 2)
            m_dicHeaders("COLUMN" & j) = j
        Next j
        m_lHeaderRowIndex = LBound(m_vData, 1) - 1
    End If
End Sub



' 초기화
Public Sub Clear()
    m_vData = Empty
    m_dicHeaders.RemoveAll
    BumpGeneration
    m_lRowCount = 0
    m_lColCount = 0
End Sub

' 세대 증가 = 전체 파생물(인덱스/RowSet/RowList) 실효. 구조 변경 경로에서만 호출.
' 값 전용 변경(Value Let, UpdateBatch, Compute, CopyColumn)은 rowID가 유효하므로
' 세대를 올리지 않고 per-column 무효화만 수행한다.
Private Sub BumpGeneration()
    m_lGeneration = m_lGeneration + 1
    InvalidateAllIndices
End Sub

Private Sub InvalidateAllIndices()
    If Not m_dicHashIndices Is Nothing Then m_dicHashIndices.RemoveAll
End Sub

Private Sub InvalidateIndexByColumnName(ByVal sColName As String)
    Dim lCol As Long

    sColName = UCase(Trim(sColName))
    If m_dicHashIndices.exists(sColName) Then m_dicHashIndices.Remove sColName

    If m_dicHeaders.exists(sColName) Then
        lCol = m_dicHeaders(sColName)
        If m_dicHashIndices.exists(lCol) Then m_dicHashIndices.Remove lCol
    End If
End Sub

' ==============================================================================
' [2] 데이터 조회 (Select/Get)
' ==============================================================================

' 2-1. 특정 셀 데이터 가져오기 (열 이름 이용)
Public Function GetValue(ByVal lRow As Long, ByVal sColName As String) As Variant
    Dim lCol As Long
    sColName = UCase(Trim(sColName))
    
    If m_dicHeaders.exists(sColName) Then
        lCol = m_dicHeaders(sColName)
        ' 배열 범위 방어
        If lRow >= LBound(m_vData, 1) And lRow <= m_lRowCount Then
            GetValue = m_vData(lRow, lCol)
        End If
    End If
End Function

' 2-2. 특정 행의 전체 레코드 가져오기 (배열 반환)
Public Function GetRecord(ByVal lRow As Long) As Variant
    Dim vRow() As Variant
    Dim j As Long, lStartCol As Long, lEndCol As Long
    
    If lRow < LBound(m_vData, 1) Or lRow > m_lRowCount Then Exit Function
    
    lStartCol = LBound(m_vData, 2)
    lEndCol = UBound(m_vData, 2)
    
    ReDim vRow(lStartCol To lEndCol)
    
    For j = lStartCol To lEndCol
        vRow(j) = m_vData(lRow, j)
    Next j
    
    GetRecord = vRow
End Function

' ==============================================================================
' [3] 인덱싱 및 검색 (Core Engine)
' ==============================================================================

' 3-1. 인덱스 생성 및 획득 (By 컬럼명)
Public Function GetHashIndex(ByVal sColName As String) As DbHashIndex
    Dim lCol As Long
    Dim idxResult As DbHashIndex
    
    sColName = UCase(Trim(sColName))
    
    If m_dicHashIndices.exists(sColName) Then
        Set idxResult = m_dicHashIndices(sColName)
        idxResult.EnsureCurrent Me ' 불변식 방어: 캐시된 인덱스는 항상 현재 세대여야 함
        Set GetHashIndex = idxResult
        Exit Function
    End If
    
    If Not m_dicHeaders.exists(sColName) Then Exit Function
    lCol = m_dicHeaders(sColName)
    
    Set idxResult = GetHashIndexByIndex(lCol)
    If Not idxResult Is Nothing Then
        Set m_dicHashIndices(sColName) = idxResult
    End If
    
    Set GetHashIndex = idxResult
End Function

' 3-1b. 인덱스 생성 및 획득 (By 숫자 인덱스)
' 빌드는 IndexBackend가 고른 구체 백엔드로, 반환은 인터페이스로
Public Function GetHashIndexByIndex(ByVal lCol As Long) As DbHashIndex
    Dim i As Long
    Dim idx As DbHashIndex

    If m_dicHashIndices.exists(lCol) Then
        Set idx = m_dicHashIndices(lCol)
        idx.EnsureCurrent Me ' 불변식 방어: 캐시된 인덱스는 항상 현재 세대여야 함
        Set GetHashIndexByIndex = idx
        Exit Function
    End If

    If lCol < 1 Or lCol > m_lColCount Then Exit Function

    Set idx = NewBoundIndex()

    ' Phase 6.8 밀봉 - 컬럼 접근은 CopyColumn1D 경유
    Dim vCol As Variant
    vCol = CopyColumn1D(lCol)
    For i = (m_lHeaderRowIndex + 1) To m_lRowCount
        idx.AddRow vCol(i), i
    Next i

    Set m_dicHashIndices(lCol) = idx
    Set GetHashIndexByIndex = idx
End Function

' [Internal] IndexBackend 선택에 따라 현재 세대에 결속된 빈 인덱스 생성
Private Function NewBoundIndex() As DbHashIndex
    If m_lIndexBackend = dbBackendCore Then
        Dim idxCore As CoreHashIndex
        Set idxCore = New CoreHashIndex
        idxCore.BindTo Me
        Set NewBoundIndex = idxCore
    Else
        Dim idxDict As DictHashIndex
        Set idxDict = New DictHashIndex
        idxDict.BindTo Me
        Set NewBoundIndex = idxDict
    End If
End Function


' 3-2a. 특정 조건에 맞는 행 번호들 찾기 (By 컬럼명) - DictHashIndex 어댑터 (호환 shim)
Public Function SelectRows(ByVal sColName As String, ByVal vValue As Variant) As Collection
    Dim idx As DbHashIndex
    Dim rows As DbRowList

    Set idx = GetHashIndex(sColName)
    If idx Is Nothing Then Set SelectRows = New Collection: Exit Function

    ' 키 정규화(Empty/CStr/대소문자 무시)는 TryGetRows가 담당
    If idx.TryGetRows(vValue, rows) Then
        Set SelectRows = rows.ToCollection
    Else
        Set SelectRows = New Collection
    End If
End Function

' 3-2b. 특정 조건에 맞는 행 번호들 찾기 (By 숫자 인덱스) - DictHashIndex 어댑터 (호환 shim)
Public Function SelectRowsByIndex(ByVal lCol As Long, ByVal vValue As Variant) As Collection
    Dim idx As DbHashIndex
    Dim rows As DbRowList

    Set idx = GetHashIndexByIndex(lCol)
    If idx Is Nothing Then Set SelectRowsByIndex = New Collection: Exit Function

    If idx.TryGetRows(vValue, rows) Then
        Set SelectRowsByIndex = rows.ToCollection
    Else
        Set SelectRowsByIndex = New Collection
    End If
End Function


' 3-3a. 특정 컬럼의 고유값 목록 추출 (By 컬럼명) - DictHashIndex 어댑터 (호환 shim)
Public Function GetUniqueKeys(ByVal sColName As String) As Variant
    Dim idx As DbHashIndex
    Set idx = GetHashIndex(sColName)

    If idx Is Nothing Then
        GetUniqueKeys = Array()
    Else
        GetUniqueKeys = KeysFromIndex(idx)
    End If
End Function

' 3-3b. 특정 컬럼의 고유값 목록 추출 (By 숫자 인덱스) - DictHashIndex 어댑터 (호환 shim)
Public Function GetUniqueKeysByIndex(ByVal lCol As Long) As Variant
    Dim idx As DbHashIndex
    Set idx = GetHashIndexByIndex(lCol)

    If idx Is Nothing Then
        GetUniqueKeysByIndex = Array()
    Else
        GetUniqueKeysByIndex = KeysFromIndex(idx)
    End If
End Function

' [Internal] 인덱스 커서 순회로 키 배열 생성 (삽입 순서 = 행 순서 유지)
Private Function KeysFromIndex(ByVal idx As DbHashIndex) As Variant
    Dim cur As DbIndexCursor
    Dim vKeys() As Variant
    Dim lPos As Long

    If idx.CountKeys = 0 Then
        KeysFromIndex = Array()
        Exit Function
    End If

    ReDim vKeys(0 To idx.CountKeys - 1)
    Set cur = idx.Cursor
    Do While cur.MoveNext
        vKeys(lPos) = cur.Key
        lPos = lPos + 1
    Loop

    KeysFromIndex = vKeys
End Function


' 3-4a. 컬럼 접근 표면 획득 (By 컬럼명)
' @desc 컬럼명 받아 세대 결속 컬럼 reader 반환, 미존재 컬럼 시 Nothing
Public Function GetColumnReader(ByVal sColName As String) As DbColumnReader
    sColName = UCase(Trim(sColName))
    If Not m_dicHeaders.exists(sColName) Then Exit Function
    Set GetColumnReader = GetColumnReaderByIndex(m_dicHeaders(sColName))
End Function

' 3-4b. 컬럼 접근 표면 획득 (By 숫자 인덱스)
' @desc 컬럼 인덱스 받아 세대 결속 컬럼 reader 반환, 범위 밖 시 Nothing
Public Function GetColumnReaderByIndex(ByVal lCol As Long) As DbColumnReader
    Dim rdr As DbColumnReader

    If lCol < 1 Or lCol > m_lColCount Then Exit Function

    Set rdr = New DbColumnReader
    rdr.BindData CopyColumn1D(lCol), ObjPtr(Me), m_lGeneration
    Set GetColumnReaderByIndex = rdr
End Function

' [Internal] 컬럼 단위 m_vData 접근의 유일한 지점 (레이아웃 밀봉 접점)
' 데이터 행 없으면 Empty 반환. 인덱스 빌드/GroupBy도 이 헬퍼를 경유한다.
Private Function CopyColumn1D(ByVal lCol As Long) As Variant
    Dim v() As Variant
    Dim i As Long

    If m_lRowCount < m_lHeaderRowIndex + 1 Then Exit Function

    ReDim v(m_lHeaderRowIndex + 1 To m_lRowCount)
    For i = m_lHeaderRowIndex + 1 To m_lRowCount
        v(i) = m_vData(i, lCol)
    Next i

    CopyColumn1D = v
End Function

' ==============================================================================
' [4] 속성 (Properties)
' ==============================================================================

' @desc 구조 변경마다 증가하는 현재 세대 번호 반환
Public Property Get Generation() As Long
    Generation = m_lGeneration
End Property

' @desc 현재 인덱스 백엔드 선택값 반환
Public Property Get IndexBackend() As DbIndexBackend
    IndexBackend = m_lIndexBackend
End Property

' @desc 인덱스 백엔드를 지정함, 변경 시 기존 인덱스 캐시 전체 실효
Public Property Let IndexBackend(ByVal vBackend As DbIndexBackend)
    If vBackend <> m_lIndexBackend Then
        m_lIndexBackend = vBackend
        InvalidateAllIndices ' 백엔드 혼합 캐시 방지
    End If
End Property

Public Property Get RowCount() As Long
    RowCount = m_lRowCount
End Property

Public Property Get ColCount() As Long
    ColCount = m_lColCount
End Property

Public Property Get Headers() As Object
    Set Headers = m_dicHeaders
End Property

Public Property Get RawArray() As Variant
    RawArray = m_vData
End Property


' [clsTable.cls] 추가 코드

' ==============================================================================
' [5] 출력 (Export)
' ==============================================================================

' 시트에 데이터 고속 출력 (Array Dump 방식)
Public Sub DumpToSheet(ByVal rngStart As Range, Optional ByVal bAutoFit As Boolean = False)
    Dim rngTarget As Range
    
    If m_lRowCount = 0 Then Exit Sub
    
    ' 1. 출력 범위 계산
    ' rngStart 셀을 기준으로 데이터 크기만큼 범위 설정
    Set rngTarget = rngStart.Resize(m_lRowCount, m_lColCount)
    
    ' 2. 기존 데이터 클리어 (선택 사항 - 안전을 위해 해당 범위만 값 덮어쓰기)
    ' rngTarget.ClearContents ' 필요 시 주석 해제
    
    ' 3. 고속 덤프 (배열을 한 번에 뿌림)
    rngTarget.Value2 = m_vData
    
    ' 4. 서식 적용 (헤더 강조)
    With rngTarget.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241) ' 연한 파랑
    End With
    
    ' 5. 컬럼 자동 맞춤
    If bAutoFit Then rngTarget.EntireColumn.AutoFit
End Sub


' [clsTable.cls] 추가 코드

' ==============================================================================
' [6] 집계 (Aggregation)
' ==============================================================================

' 그룹화 및 집계 함수
' sGroupCols: 그룹핑할 컬럼명 (콤마로 구분, 예: "Dept, Team")
' sAggList: 집계할 컬럼과 함수 (콤마로 구분, 예: "Salary:SUM, Age:AVG, ID:COUNT")
Public Function GroupBy(ByVal sGroupCols As String, ByVal sAggList As String) As clsTable
    
    Dim dicAccumulator As Object ' Key -> Array(Running Stats)
    Dim dicGroupParts As Object  ' Key -> 정규화된 그룹 값 파트 (첫 등장 표기, 렌더용)
    Dim arrParts() As String
    Dim vRow As Variant
    Dim vStats As Variant
    Dim sKey As String, sVal As String
    
    ' 파싱 변수
    Dim arrGroupCols() As String, arrGroupIndices() As Long
    Dim arrAggItems() As String, arrAggIndices() As Long, arrAggFuncs() As String
    Dim lGroupCount As Long, lAggCount As Long
    
    Dim i As Long, j As Long, R As Long
    Dim vCell As Variant
    Dim bSkipRow As Boolean

    ' 초기화 - 키 의미론은 인덱스 계약(modIndexKey)과 동일: 대소문자 무시
    Set dicAccumulator = CreateObject("Scripting.Dictionary")
    dicAccumulator.CompareMode = vbTextCompare
    Set dicGroupParts = CreateObject("Scripting.Dictionary")
    dicGroupParts.CompareMode = vbTextCompare
    arrGroupCols = Split(sGroupCols, ",")
    arrAggItems = Split(sAggList, ",")
    lGroupCount = UBound(arrGroupCols) + 1
    lAggCount = UBound(arrAggItems) + 1
    
    ' 1. [Parsing] 컬럼 인덱스 및 함수 매핑
    ReDim arrGroupIndices(0 To lGroupCount - 1)
    For i = 0 To lGroupCount - 1
        sVal = Trim(arrGroupCols(i))
        If m_dicHeaders.exists(UCase(sVal)) Then
            arrGroupIndices(i) = m_dicHeaders(UCase(sVal))
        Else
            Err.Raise 9999, , "Group Column Not Found: " & sVal
        End If
    Next i
    
    ReDim arrAggIndices(0 To lAggCount - 1)
    ReDim arrAggFuncs(0 To lAggCount - 1)
    
    Dim arrTemp() As String
    For i = 0 To lAggCount - 1
        sVal = Trim(arrAggItems(i))
        ' "ColName:Func" 형태 분리
        If InStr(sVal, ":") > 0 Then
            arrTemp = Split(sVal, ":")
            sVal = Trim(arrTemp(0))
            ' 화이트리스트 검증 (audit F2, V1): 미지 함수는 즉시 raise - 이전엔 누적
            ' 분기 미매칭 + 렌더 fallback으로 "조용한 빈 컬럼"이 나왔다 (Agg Column
            ' Not Found의 9999와 오류 처리 대칭)
            arrAggFuncs(i) = modOperators.NormalizeAggFunc(arrTemp(1))
        Else
            ' 함수 명시 없으면 기본 COUNT로 처리
            arrAggFuncs(i) = "COUNT"
        End If
        
        If m_dicHeaders.exists(UCase(sVal)) Then
            arrAggIndices(i) = m_dicHeaders(UCase(sVal))
        Else
            Err.Raise 9999, , "Agg Column Not Found: " & sVal
        End If
    Next i
    
    ' 2-0. [Hoisting] 컬럼 접근 밀봉 (Phase 6.8) - 대상 컬럼만 1D 사본으로 물질화
    Dim vGroupCols() As Variant
    Dim vAggCols() As Variant
    ReDim vAggCols(0 To lAggCount - 1)
    For i = 0 To lAggCount - 1
        vAggCols(i) = CopyColumn1D(arrAggIndices(i))
    Next i

    Dim tResult As New clsTable
    Dim vResultData As Variant
    Dim lResultRows As Long, lResultCols As Long
    Dim vKeys As Variant

    lResultCols = lGroupCount + lAggCount

    ' 2-S. [R2 엔진 경로] 단일 그룹 컬럼: 해시 인덱스(키->RowList)가 이미 그룹 파티션.
    ' 커서 순회로 그룹별 RowList만 읽어 전체 행 재스캔과 로컬 누적기를 제거한다.
    ' 키 의미론(vbTextCompare/Empty->""/Error skip)은 인덱스 계약이 그대로 보장한다.
    If lGroupCount = 1 Then
        Dim idx As DbHashIndex
        Dim cur As DbIndexCursor
        Dim bufRows() As Long ' 그룹 span 재사용 버퍼 (키당 객체 생성 없음)
        Dim cntRows As Long
        Dim lRowId As Long, k As Long

        Set idx = GetHashIndexByIndex(arrGroupIndices(0))
        lResultRows = idx.CountKeys
        ReDim vResultData(1 To lResultRows + 1, 1 To lResultCols)

        vResultData(1, 1) = arrGroupCols(0)
        For i = 0 To lAggCount - 1
            vResultData(1, lGroupCount + i + 1) = Trim(Split(arrAggItems(i), ":")(0)) & "_" & arrAggFuncs(i)
        Next i

        Set cur = idx.Cursor
        R = 1
        Do While cur.MoveNext
            R = R + 1
            vResultData(R, 1) = cur.Key
            cur.CopyRows bufRows, cntRows

            ReDim vStats(0 To (lAggCount * 2) - 1)
            For j = 0 To lAggCount - 1
                If arrAggFuncs(j) = "MIN" Then vStats(j * 2) = Empty
                If arrAggFuncs(j) = "MAX" Then vStats(j * 2) = Empty
            Next j

            For k = 1 To cntRows
                lRowId = bufRows(k)
                For j = 0 To lAggCount - 1
                    vCell = vAggCols(j)(lRowId)
                    ' 집계 본체는 복합 키 경로(2-3)와 동일 계약 - R3에서 공용화 예정
                    If Not IsEmpty(vCell) Then
                        ' 집계 계약(D2): Empty=NULL/Error=skip - 둘 다 불참(Empty는 위 가드).
                        ' COUNT=Error 제외한 실값 수, AVG 분모=숫자로 가산된 셀 수(희석 방지),
                        ' MIN/MAX=Error 제외(raw 비교 크래시 차단).
                        Select Case arrAggFuncs(j)
                            Case "SUM"
                                If IsNumeric(vCell) Then vStats(j * 2) = vStats(j * 2) + vCell
                            Case "AVG"
                                If IsNumeric(vCell) Then
                                    vStats(j * 2) = vStats(j * 2) + vCell
                                    vStats(j * 2 + 1) = vStats(j * 2 + 1) + 1
                                End If
                            Case "COUNT"
                                If Not IsError(vCell) Then vStats(j * 2 + 1) = vStats(j * 2 + 1) + 1
                            Case "MAX"
                                If Not IsError(vCell) Then
                                    If IsEmpty(vStats(j * 2)) Then
                                        vStats(j * 2) = vCell
                                    ElseIf vCell > vStats(j * 2) Then
                                        vStats(j * 2) = vCell
                                    End If
                                End If
                            Case "MIN"
                                If Not IsError(vCell) Then
                                    If IsEmpty(vStats(j * 2)) Then
                                        vStats(j * 2) = vCell
                                    ElseIf vCell < vStats(j * 2) Then
                                        vStats(j * 2) = vCell
                                    End If
                                End If
                        End Select
                    End If
                Next j
            Next k

            For j = 0 To lAggCount - 1
                Select Case arrAggFuncs(j)
                    Case "AVG"
                        If vStats(j * 2 + 1) > 0 Then
                            vResultData(R, lGroupCount + j + 1) = vStats(j * 2) / vStats(j * 2 + 1)
                        Else
                            vResultData(R, lGroupCount + j + 1) = 0
                        End If
                    Case "COUNT"
                        vResultData(R, lGroupCount + j + 1) = vStats(j * 2 + 1)
                    Case Else ' SUM, MAX, MIN
                        vResultData(R, lGroupCount + j + 1) = vStats(j * 2)
                End Select
            Next j
        Loop

        tResult.AdoptArray vResultData
        Set GroupBy = tResult
        Exit Function
    End If

    ' 복합 키 경로 (R3에서 encode helper 도입과 엔진화 여부 결정 예정)
    ReDim vGroupCols(0 To lGroupCount - 1)
    For i = 0 To lGroupCount - 1
        vGroupCols(i) = CopyColumn1D(arrGroupIndices(i))
    Next i

    ' 2. [Aggregation] 데이터 순회 및 집계 (Main Loop)
    ' Value 배열 구조: [Agg1_Val, Agg1_Count, Agg2_Val, Agg2_Count ...]
    ' AVG 계산을 위해 모든 항목에 대해 Count를 별도로 관리합니다.

    For R = (m_lHeaderRowIndex + 1) To m_lRowCount
        ' 2-1. 복합 키(Composite Key) 생성 - 길이 접두 인코딩(modIndexKey, R3)으로 구분자 충돌 차단
        ' 인덱스 계약 준수: Empty->"", Error 키 행은 그룹핑에서 제외 (인덱스 AddRow skip과 동치)
        sKey = ""
        bSkipRow = False
        For i = 0 To lGroupCount - 1
            If IsError(vGroupCols(i)(R)) Then
                bSkipRow = True
                Exit For
            End If
            sKey = sKey & modIndexKey.EncodeCompositeKeyPart(vGroupCols(i)(R)) & "|"
        Next i

        If Not bSkipRow Then
            ' 2-2. 누적기 초기화 (첫 발견 시) - 렌더용 그룹 값 파트도 첫 등장 표기로 보존
            If Not dicAccumulator.exists(sKey) Then
                ReDim vStats(0 To (lAggCount * 2) - 1) ' [Val, Cnt, Val, Cnt...]
                ' MIN/MAX 초기화 로직 필요시 추가 (여기선 간단히 0/Empty 처리)
                For j = 0 To lAggCount - 1
                    If arrAggFuncs(j) = "MIN" Then vStats(j * 2) = Empty
                    If arrAggFuncs(j) = "MAX" Then vStats(j * 2) = Empty
                Next j
                dicAccumulator(sKey) = vStats

                ReDim arrParts(0 To lGroupCount - 1)
                For i = 0 To lGroupCount - 1
                    arrParts(i) = modIndexKey.NormalizeIndexKey(vGroupCols(i)(R))
                Next i
                dicGroupParts(sKey) = arrParts
            End If

            ' 2-3. 상태 업데이트 (Accumulate)
            vStats = dicAccumulator(sKey)

            For j = 0 To lAggCount - 1
                vCell = vAggCols(j)(R)

                ' 집계 계약(D2): Empty=NULL(제외), Error=skip(어느 집계에도 불참).
                ' COUNT=Error 제외한 실값 수, AVG 분모=숫자로 가산된 셀 수, MIN/MAX=Error 제외.
                If Not IsEmpty(vCell) Then
                    Select Case arrAggFuncs(j)
                        Case "SUM"
                            If IsNumeric(vCell) Then vStats(j * 2) = vStats(j * 2) + vCell
                        Case "AVG"
                            If IsNumeric(vCell) Then
                                vStats(j * 2) = vStats(j * 2) + vCell
                                vStats(j * 2 + 1) = vStats(j * 2 + 1) + 1
                            End If
                        Case "COUNT"
                            If Not IsError(vCell) Then vStats(j * 2 + 1) = vStats(j * 2 + 1) + 1
                        Case "MAX"
                            If Not IsError(vCell) Then
                                If IsEmpty(vStats(j * 2)) Then
                                    vStats(j * 2) = vCell
                                ElseIf vCell > vStats(j * 2) Then
                                    vStats(j * 2) = vCell
                                End If
                            End If
                        Case "MIN"
                            If Not IsError(vCell) Then
                                If IsEmpty(vStats(j * 2)) Then
                                    vStats(j * 2) = vCell
                                ElseIf vCell < vStats(j * 2) Then
                                    vStats(j * 2) = vCell
                                End If
                            End If
                    End Select
                End If
            Next j

            dicAccumulator(sKey) = vStats ' Value타입이므로 다시 넣어줘야 함 (Variant Array 특성)
        End If
    Next R
    
    ' 3. [Render] 결과 테이블 생성
    lResultRows = dicAccumulator.Count
    ReDim vResultData(1 To lResultRows + 1, 1 To lResultCols)
    
    ' 3-1. 헤더 작성
    For i = 0 To lGroupCount - 1
        vResultData(1, i + 1) = arrGroupCols(i)
    Next i
    For i = 0 To lAggCount - 1
        ' 헤더명: "Salary_SUM" 형태로 자동 생성
        vResultData(1, lGroupCount + i + 1) = Trim(Split(arrAggItems(i), ":")(0)) & "_" & arrAggFuncs(i)
    Next i
    
    ' 3-2. 데이터 채우기
    vKeys = dicAccumulator.Keys
    For R = 0 To lResultRows - 1
        sKey = vKeys(R)
        vStats = dicAccumulator(sKey)
        
        ' 그룹 컬럼 복원 - 키 재분해(Split) 대신 첫 등장 시 보존한 파트 사용 (인코딩 무관)
        arrTemp = dicGroupParts(sKey)
        For i = 0 To lGroupCount - 1
            vResultData(R + 2, i + 1) = arrTemp(i)
        Next i
        
        ' 집계 결과 계산 및 입력
        For i = 0 To lAggCount - 1
            Select Case arrAggFuncs(i)
                Case "AVG"
                    ' SUM / COUNT
                    If vStats(i * 2 + 1) > 0 Then
                        vResultData(R + 2, lGroupCount + i + 1) = vStats(i * 2) / vStats(i * 2 + 1)
                    Else
                        vResultData(R + 2, lGroupCount + i + 1) = 0
                    End If
                Case "COUNT"
                    ' COUNT 값 반환
                    vResultData(R + 2, lGroupCount + i + 1) = vStats(i * 2 + 1)
                Case Else ' SUM, MAX, MIN
                    vResultData(R + 2, lGroupCount + i + 1) = vStats(i * 2)
            End Select
        Next i
    Next R
    
    tResult.AdoptArray vResultData
    Set GroupBy = tResult

End Function

' [clsTable.cls] 추가 코드

' ==============================================================================
' [7] 정렬 (Sorting)
' ==============================================================================

' 특정 컬럼 기준 정렬 (QuickSort Wrapper)
Public Function OrderBy(ByVal sColName As String, Optional ByVal bDescending As Boolean = False) As clsTable
    Dim lColIdx As Long
    Dim arrIndices() As Long
    Dim i As Long
    
    ' 1. 컬럼 확인
    sColName = UCase(Trim(sColName))
    If Not m_dicHeaders.exists(sColName) Then
        Err.Raise 9999, , "Sort Column Not Found: " & sColName
    End If
    lColIdx = m_dicHeaders(sColName)
    
    ' 2. 인덱스 배열 생성 (헤더 제외 데이터 행 번호만)
    ' m_vData의 실제 행 인덱스를 담습니다.
    ReDim arrIndices(0 To m_lRowCount - (m_lHeaderRowIndex + 1))
    For i = 0 To UBound(arrIndices)
        arrIndices(i) = m_lHeaderRowIndex + 1 + i
    Next i
    
    ' 3. 퀵 정렬 실행 (인덱스만 정렬)
    If UBound(arrIndices) >= 0 Then
        Call QuickSort(arrIndices, 0, UBound(arrIndices), lColIdx, bDescending)
    End If
    
    ' 4. 데이터 재구성 (Rebuild)
    Call RebuildData(arrIndices)

    Set OrderBy = Me
End Function
' ==============================================================================
' [8] 비즈니스 로직 (Business Logic) - 상태 변경(Mutation) 메서드
' 변형 메서드는 Me를 반환해 체인을 유지한다 (audit F16, in-place 의미론은 불변)
' ==============================================================================

' 8-1. Upsert (Merge): 다른 테이블의 데이터를 키 기준으로 병합 (Insert + Update)
' 설명: 키가 같으면 덮어쓰고(Update), 없으면 추가(Insert)합니다.
' 전략: Update는 In-place로 즉시 처리, Insert는 모아서 배치로 처리 (배열 재할당 비용 최소화)
Public Function UpsertFrom(ByVal tSource As clsTable, ByVal sKeyCol As String) As clsTable
    Dim idxKey As DbHashIndex
    Dim vSourceData As Variant
    Dim vKey As Variant
    Dim lSourceRow As Long, lTargetRow As Long
    Dim i As Long, j As Long
    Dim bUpdated As Boolean
    
    Dim colNewRows As Collection ' 신규 행 임시 저장소
    Dim vNewRow As Variant
    
    ' 1. 컬럼 매핑 (Source Index -> Target Index)
    ' 헤더 이름이 일치하는 컬럼만 업데이트 대상이 됩니다.
    Dim dicColMap As Object
    Dim lSourceColIdx As Long, lTargetColIdx As Long
    Dim sColName As String
    
    Set dicColMap = CreateObject("Scripting.Dictionary")

    vSourceData = tSource.RawArray ' 1회 물질화 (RawArray는 호출마다 전체 복사)

    ' Source의 헤더를 순회하며 Target에 존재하는지 확인
    For i = 1 To tSource.ColCount
        sColName = UCase(Trim(vSourceData(1, i)))
        If m_dicHeaders.exists(sColName) Then
            dicColMap(i) = m_dicHeaders(sColName)
        End If
    Next i
    
    ' 키 컬럼 인덱스 확보
    Dim lKeyColSource As Long
    If tSource.Headers.exists(UCase(sKeyCol)) Then
        lKeyColSource = tSource.Headers(UCase(sKeyCol))
    Else
        Err.Raise 9999, , "Upsert Error: Source table missing key column [" & sKeyCol & "]"
    End If
    
    ' 2. 타겟 테이블 인덱싱 (고속 검색 준비)
    Set idxKey = Me.GetHashIndex(sKeyCol)
    If idxKey Is Nothing Then Err.Raise 9999, , "Upsert Error: Target table missing key column"

    Set colNewRows = New Collection

    ' 3. Source 데이터 순회 및 분류 (Update vs Insert)
    ' 키 정규화(Empty/CStr/vbTextCompare)는 TryGetFirstRowId가 담당
    For lSourceRow = 2 To tSource.RowCount
        vKey = vSourceData(lSourceRow, lKeyColSource)

        If idxKey.TryGetFirstRowId(vKey, lTargetRow) Then
            ' [UPDATE] 키 존재 -> 기존 행 덮어쓰기 (In-place)
            ' 주의: 중복 키가 있다면 첫 번째 매칭 행만 업데이트

            For Each vKey In dicColMap.Keys
                lSourceColIdx = vKey
                lTargetColIdx = dicColMap(vKey)
                m_vData(lTargetRow, lTargetColIdx) = vSourceData(lSourceRow, lSourceColIdx)
            Next vKey
            bUpdated = True
        Else
            ' [INSERT] 키 미존재 -> 신규 행 임시 저장
            ' 전체 컬럼 크기만큼 배열 확보
            ReDim vNewRow(1 To m_lColCount)
            For Each vKey In dicColMap.Keys
                lSourceColIdx = vKey
                lTargetColIdx = dicColMap(vKey)
                vNewRow(lTargetColIdx) = vSourceData(lSourceRow, lSourceColIdx)
            Next vKey
            colNewRows.Add vNewRow
        End If
    Next lSourceRow
    
    ' 4. 신규 데이터 병합 (Array Resize)
    ' VBA 2차원 배열은 첫 번째 차원(Row) 확장이 불가능하므로, 새 배열을 만들어 이사해야 함
    If colNewRows.Count > 0 Then
        Dim lOldCount As Long, lNewTotal As Long
        Dim vNewData As Variant
        
        lOldCount = m_lRowCount
        lNewTotal = lOldCount + colNewRows.Count
        
        ReDim vNewData(1 To lNewTotal, 1 To m_lColCount)
        
        ' (1) 기존 데이터 복사 (헤더 포함)
        For i = 1 To lOldCount
            For j = 1 To m_lColCount
                vNewData(i, j) = m_vData(i, j)
            Next j
        Next i
        
        ' (2) 신규 데이터 추가
        i = lOldCount + 1
        For Each vNewRow In colNewRows
            For j = 1 To m_lColCount
                vNewData(i, j) = vNewRow(j)
            Next j
            i = i + 1
        Next vNewRow
        
        ' (3) 교체 및 상태 갱신
        m_vData = vNewData
        m_lRowCount = lNewTotal
        
        ' 중요: 행 번호가 바뀌거나 추가되었으므로 세대 증가 (전체 파생물 실효)
        BumpGeneration
    ElseIf bUpdated Then
        BumpGeneration
    End If

    Set UpsertFrom = Me
End Function

' 8-2. Batch Update: 조건에 맞는 행의 특정 값을 일괄 변경
' 예: UpdateBatch "Status", "Inactive", "LastLogin", "< 2023-01-01" (현재는 Equal 조건만 지원)
Public Function UpdateBatch(ByVal sSetCol As String, ByVal vSetVal As Variant, _
                            ByVal sWhereCol As String, ByVal vWhereVal As Variant) As clsTable
    
    Dim colTargetRows As Collection
    Dim vRowID As Variant
    Dim lSetColIdx As Long
    
    ' 컬럼 존재 확인
    sSetCol = UCase(Trim(sSetCol))
    If Not m_dicHeaders.exists(sSetCol) Then Err.Raise 9999, , "Column Not Found: " & sSetCol
    lSetColIdx = m_dicHeaders(sSetCol)
    
    ' 조건 검색 (기존 SelectRows 활용 - 인덱스 기반 고속 검색)
    Set colTargetRows = Me.SelectRows(sWhereCol, vWhereVal)
    
    ' 데이터 수정
    If colTargetRows.Count > 0 Then
        For Each vRowID In colTargetRows
            m_vData(CLng(vRowID), lSetColIdx) = vSetVal
        Next vRowID
        
        ' 해당 컬럼의 데이터가 변했으므로 관련 인덱스 캐시를 제거
        InvalidateIndexByColumnName sSetCol
    End If

    Set UpdateBatch = Me
End Function

' 8-3. Batch Delete: 조건에 맞는 행을 일괄 삭제
' 전략: 삭제 대상이 아닌 행만 골라내어 새 배열로 옮기는 Rebuild 방식 (Copy Alive)
Public Function DeleteBatch(ByVal sWhereCol As String, ByVal vWhereVal As Variant) As clsTable
    Dim colDeleteRows As Collection
    Dim dicDeleteMap As Object
    Dim vRowID As Variant
    Dim vNewData As Variant
    Dim lNewRowCount As Long
    Dim i As Long, j As Long, lCurrentRow As Long
    
    ' 삭제 대상 검색 (0건 매칭은 데이터 수준 - 무동작이 계약)
    Set colDeleteRows = Me.SelectRows(sWhereCol, vWhereVal)
    If colDeleteRows.Count = 0 Then Set DeleteBatch = Me: Exit Function
    
    ' 빠른 조회를 위해 Dictionary로 변환 (O(1) Lookup)
    Set dicDeleteMap = CreateObject("Scripting.Dictionary")
    For Each vRowID In colDeleteRows
        dicDeleteMap(vRowID) = Empty
    Next vRowID
    
    ' 새 배열 크기 계산
    lNewRowCount = m_lRowCount - colDeleteRows.Count
    If lNewRowCount < 1 Then lNewRowCount = 1 ' 최소 헤더 1행 유지
    
    ReDim vNewData(1 To lNewRowCount, 1 To m_lColCount)
    
    ' 데이터 이관
    lCurrentRow = 1
    
    ' (1) 헤더는 무조건 복사 (Row 1)
    For j = 1 To m_lColCount
        vNewData(1, j) = m_vData(1, j)
    Next j
    lCurrentRow = 2
    
    ' (2) 데이터 복사 (삭제 대상 제외)
    For i = 2 To m_lRowCount
        If Not dicDeleteMap.exists(i) Then
            For j = 1 To m_lColCount
                vNewData(lCurrentRow, j) = m_vData(i, j)
            Next j
            lCurrentRow = lCurrentRow + 1
        End If
    Next i
    
    ' 교체 및 상태 갱신
    m_vData = vNewData
    m_lRowCount = lNewRowCount
    
    ' 행 번호가 완전히 바뀌었으므로 세대 증가 (전체 파생물 실효)
    BumpGeneration

    Set DeleteBatch = Me
End Function

' ==============================================================================
' [9] 데이터 조작 (Data Manipulation) - v1.3 추가 기능
' ==============================================================================

' 9-1-1. 값 직접 접근 (읽기/쓰기 겸용)
' 사용법: t.Value(1, "Status") = "Completed"
'         Debug.Print t.Value(1, "Status")
Public Property Let Value(ByVal lRow As Long, ByVal sColName As String, ByVal vVal As Variant)
    Dim lCol As Long
    sColName = UCase(Trim(sColName))
    
    If m_dicHeaders.exists(sColName) Then
        lCol = m_dicHeaders(sColName)
        ' 행 범위 체크
        If lRow >= (m_lHeaderRowIndex + 1) And lRow <= m_lRowCount Then
            m_vData(lRow, lCol) = vVal
            InvalidateIndexByColumnName sColName
        End If
    Else
        ' 미지 컬럼 쓰기는 문장 수준 오류 - fail-fast (audit 추보 F24, V2).
        ' 이전 Debug.Print 경고는 헤드리스에서 보이지 않는 조용한 무시였다.
        Err.Raise 9999, "clsTable.Value(Let)", "Column not found: " & sColName
    End If
End Property
' 9-1-2 clsTable에 추가해두면 좋은 고속 복사 메서드
Public Function CopyColumn(ByVal sSourceCol As String, ByVal sTargetCol As String) As clsTable
    Dim idxSrc As Long, idxDst As Long
    Dim i As Long

    ' 소스 컬럼 오타는 문장 수준 오류 - fail-fast (audit 추보 F23, 이전엔 조용한 무동작)
    If Not m_dicHeaders.exists(UCase(sSourceCol)) Then Err.Raise 9999, "clsTable.CopyColumn", "Column not found: " & sSourceCol
    idxSrc = m_dicHeaders(UCase(sSourceCol))
    
    ' 타겟 컬럼 없으면 생성, 있으면 인덱스 확보
    If Not m_dicHeaders.exists(UCase(sTargetCol)) Then AddColumn sTargetCol
    idxDst = m_dicHeaders(UCase(sTargetCol))
    
    ' 배열 직접 접근 (가장 빠름)
    For i = (m_lHeaderRowIndex + 1) To m_lRowCount
        m_vData(i, idxDst) = m_vData(i, idxSrc)
    Next i
    
    InvalidateIndexByColumnName sTargetCol

    Set CopyColumn = Me
End Function

' 9-2-1. 새로운 컬럼 추가 (리포트 생성 시 필수)
' sColName: 추가할 헤더명
' vInitialValue: (선택) 초기값. 생략 시 Empty.
Public Function AddColumn(ByVal sColName As String, Optional ByVal vInitialValue As Variant) As clsTable
    Dim i As Long

    ' 1. 중복 컬럼명 방지 (멱등 ensure 계약 - 이미 존재하면 무시, 문서화는 C4)
    sColName = Trim(sColName)
    If m_dicHeaders.exists(UCase(sColName)) Then Set AddColumn = Me: Exit Function
    
    ' 2. 컬럼 카운트 증가
    m_lColCount = m_lColCount + 1
    
    ' 3. 배열 확장 (Preserve는 마지막 차원만 변경 가능하므로 Row, Col 구조에서 안전함)
    ReDim Preserve m_vData(1 To m_lRowCount, 1 To m_lColCount)
    
    ' 4. 헤더 등록
    ' 헤더 행에 새 컬럼 이름 입력
    m_vData(m_lHeaderRowIndex, m_lColCount) = sColName
    m_dicHeaders(UCase(sColName)) = m_lColCount
    
    ' 5. 초기값 채우기 (필요한 경우만)
    If Not IsMissing(vInitialValue) Then
        For i = (m_lHeaderRowIndex + 1) To m_lRowCount
            m_vData(i, m_lColCount) = vInitialValue
        Next i
    End If

    BumpGeneration

    Set AddColumn = Me
End Function
' 9-2-2. 기존 컬럼 삭제
' sColName: 삭제할 헤더명
Public Function RemoveColumn(ByVal sColName As String) As clsTable
    Dim lTargetIdx As Long
    Dim i As Long, j As Long, k As Long
    Dim vNewData() As Variant

    sColName = Trim(sColName)
    ' 멱등 drop 계약 - 없으면 무시 (문서화는 C4)
    If Not m_dicHeaders.exists(UCase(sColName)) Then Set RemoveColumn = Me: Exit Function
    
    ' 1. 삭제할 컬럼의 인덱스 확인
    lTargetIdx = m_dicHeaders(UCase(sColName))
    
    ' 2. 새로운 배열 선언 (컬럼 수가 1개 줄어듦)
    ReDim vNewData(1 To m_lRowCount, 1 To m_lColCount - 1)
    
    ' 3. 데이터 복사 (삭제 대상 컬럼만 제외하고 이동)
    k = 0
    For j = 1 To m_lColCount
        If j <> lTargetIdx Then
            k = k + 1 ' 새 배열의 컬럼 인덱스
            For i = 1 To m_lRowCount
                vNewData(i, k) = m_vData(i, j)
            Next i
        End If
    Next j
    
    ' 4. 클래스 내부 변수 갱신
    m_vData = vNewData ' 배열 교체
    m_lColCount = m_lColCount - 1
    
    ' 5. 헤더 딕셔너리 재구성 (인덱스가 하나씩 밀리므로 새로 고침)
    Set m_dicHeaders = CreateObject("Scripting.Dictionary")
    For j = 1 To m_lColCount
        m_dicHeaders(UCase(Trim(m_vData(m_lHeaderRowIndex, j)))) = j
    Next j

    BumpGeneration

    Set RemoveColumn = Me
End Function

' ==============================================================================
' [기능 개선] 특정 열만 남기기 (Select Columns) - ParamArray 적용
' 설명: 입력받은 컬럼만 순서대로 남기고 나머지는 제거합니다. (순서 변경 가능)
' 사용법 1: t.SelectColumns "Name", "Age", "ID"      (가장 자연스러움)
' 사용법 2: t.SelectColumns "Name, Age, ID"          (기존 호환)
' 사용법 3: t.SelectColumns Array("Name", "Age")     (배열 전달 가능)
' ==============================================================================
' 목록 인자 플랫 파싱 (개별 인자/csv 한 줄/Array 허용) - SelectColumns/AddColumns/
' RemoveColumns 공용. 쿼리 실행 경로가 아닌 인자 파싱 표면이라 Collection 사용 허용.
Private Function FlattenColumnArgs(ByVal vArgs As Variant) As Collection
    Dim vItem As Variant, vSubItem As Variant
    Dim arrTemp As Variant
    Dim sColName As String

    Set FlattenColumnArgs = New Collection
    If UBound(vArgs) < LBound(vArgs) Then Exit Function

    For Each vItem In vArgs
        If IsArray(vItem) Then
            ' Case: Array("A", "B")가 넘어온 경우
            For Each vSubItem In vItem
                If Not IsError(vSubItem) And Not IsEmpty(vSubItem) Then
                    sColName = Trim(CStr(vSubItem))
                    If Len(sColName) > 0 Then FlattenColumnArgs.Add sColName
                End If
            Next vSubItem
        ElseIf Not IsError(vItem) And Not IsEmpty(vItem) Then
            ' Case: "A, B, C" 또는 "A" 단일 값이 넘어온 경우
            If InStr(CStr(vItem), ",") > 0 Then
                arrTemp = Split(CStr(vItem), ",")
                For Each vSubItem In arrTemp
                    sColName = Trim(CStr(vSubItem))
                    If Len(sColName) > 0 Then FlattenColumnArgs.Add sColName
                Next vSubItem
            Else
                sColName = Trim(CStr(vItem))
                If Len(sColName) > 0 Then FlattenColumnArgs.Add sColName
            End If
        End If
    Next vItem
End Function

' @desc 컬럼 여러 개 추가 후 자기 자신 반환, 개별 인자/csv/Array 허용 (멱등 ensure)
Public Function AddColumns(ParamArray vCols() As Variant) As clsTable
    Dim colNames As Collection
    Dim vName As Variant
    Set colNames = FlattenColumnArgs(vCols)
    For Each vName In colNames
        AddColumn CStr(vName)
    Next vName
    Set AddColumns = Me
End Function

' @desc 컬럼 여러 개 삭제 후 자기 자신 반환, 개별 인자/csv/Array 허용 (멱등 drop)
Public Function RemoveColumns(ParamArray vCols() As Variant) As clsTable
    Dim colNames As Collection
    Dim vName As Variant
    Set colNames = FlattenColumnArgs(vCols)
    For Each vName In colNames
        RemoveColumn CStr(vName)
    Next vName
    Set RemoveColumns = Me
End Function

Public Function SelectColumns(ParamArray vCols() As Variant) As clsTable
    Dim colFinalList As Collection

    Dim vNewData() As Variant
    Dim arrOldIndices() As Long
    Dim i As Long, j As Long
    Dim lNewColCount As Long
    Dim sColName As String

    ' 1. 인자 파싱 (ParamArray -> Flat Collection, 공용 헬퍼)
    Set colFinalList = FlattenColumnArgs(vCols)

    lNewColCount = colFinalList.Count
    If lNewColCount = 0 Then Set SelectColumns = Me: Exit Function
    
    ' 2. 유효성 검사 및 구(Old) 인덱스 매핑
    ReDim arrOldIndices(1 To lNewColCount)
    
    For j = 1 To lNewColCount
        sColName = colFinalList(j)
        ' 존재하지 않는 컬럼 요청 시 에러 발생
        If Not m_dicHeaders.exists(UCase(sColName)) Then
            Err.Raise 9999, "clsTable.SelectColumns", "Column not found: [" & sColName & "]"
        End If
        arrOldIndices(j) = m_dicHeaders(UCase(sColName))
    Next j
    
    ' 3. 새 배열 생성 및 데이터 고속 복사
    ReDim vNewData(1 To m_lRowCount, 1 To lNewColCount)
    
    For i = 1 To m_lRowCount
        For j = 1 To lNewColCount
            vNewData(i, j) = m_vData(i, arrOldIndices(j))
        Next j
    Next i
    
    ' 4. 상태 갱신 (Swap & Clear Cache)
    m_vData = vNewData
    m_lColCount = lNewColCount
    
    ' 헤더 딕셔너리 재구축
    m_dicHeaders.RemoveAll
    For j = 1 To m_lColCount
        sColName = Trim(CStr(m_vData(m_lHeaderRowIndex, j)))
        m_dicHeaders(UCase(sColName)) = j
    Next j

    BumpGeneration

    Set SelectColumns = Me
End Function


' 9-3. 컬럼 간/상수 연산 (Loop 없는 수식 대체)
' 사용법: t.Compute "Total", "Qty", "*", "Price" (컬럼 x 컬럼)
'         t.Compute "Tax", "Amt", "*", 0.1     (컬럼 x 숫자 상수)
' 지원 연산자: +, -, *, /, & (문자열 결합)
' 피연산자 구분 규칙: String이면 컬럼명(부재 시 9999), 그 외 타입이면 숫자 상수.
' 문자열 리터럴 상수는 미지원 (String=컬럼명 규칙과 충돌 - V9 공백으로 문서화).
' @desc 컬럼/숫자 상수 피연산자 연산해 타겟 컬럼 기록, 문장 오류 시 9999
Public Function Compute(ByVal sTargetCol As String, ByVal vOperand1 As Variant, ByVal sOp As String, ByVal vOperand2 As Variant) As clsTable
    Dim idxTarget As Long, idx1 As Long, idx2 As Long
    Dim bConst1 As Boolean, bConst2 As Boolean
    Dim vConst1 As Variant, vConst2 As Variant
    Dim i As Long
    Dim v1 As Variant, v2 As Variant, vRes As Variant

    ' 문장 수준 검증 선행 (audit F3, V1/V2): 연산자/피연산자 오류는 구조 변경(타겟
    ' 컬럼 자동 생성) 전에 raise - 이전엔 타겟만 만들고 조용히 무동작하는 반쪽 실행.
    sOp = modOperators.NormalizeComputeOp(sOp)
    ResolveComputeOperand vOperand1, bConst1, idx1, vConst1
    ResolveComputeOperand vOperand2, bConst2, idx2, vConst2

    ' 상수 0 나눗셈은 문장에 쓰인 개발자 오류 - 문장 수준 raise
    ' (컬럼 셀의 0은 기존 데이터 수준 관용 유지: 해당 행만 0 기입)
    If sOp = "/" And bConst2 Then
        If CDbl(vConst2) = 0 Then Err.Raise 9999, "clsTable.Compute", "Constant divisor must not be zero"
    End If

    ' 타겟 컬럼 확보 (없으면 자동 생성)
    If Not m_dicHeaders.exists(UCase(sTargetCol)) Then AddColumn sTargetCol
    idxTarget = m_dicHeaders(UCase(sTargetCol))

    ' 고속 배열 연산
    For i = (m_lHeaderRowIndex + 1) To m_lRowCount
        If bConst1 Then v1 = vConst1 Else v1 = m_vData(i, idx1)
        If bConst2 Then v2 = vConst2 Else v2 = m_vData(i, idx2)

        ' 행 데이터 수준 오류(비숫자 셀 등)는 계약된 관용 - 해당 셀만 미기입.
        ' 문장 수준 오류(위 raise)와 구분되는 데이터 수준 skip이다.
        On Error Resume Next
        Select Case sOp
            Case "+": m_vData(i, idxTarget) = v1 + v2
            Case "-": m_vData(i, idxTarget) = v1 - v2
            Case "*": m_vData(i, idxTarget) = v1 * v2
            Case "/": If v2 <> 0 Then m_vData(i, idxTarget) = v1 / v2 Else m_vData(i, idxTarget) = 0
            Case "&": m_vData(i, idxTarget) = CStr(v1) & CStr(v2)
        End Select
        On Error GoTo 0
    Next i

    InvalidateIndexByColumnName sTargetCol

    Set Compute = Me
End Function

' 피연산자 해석 (Compute 전용): String은 컬럼명(부재 시 9999), 그 외는 숫자 상수.
' IsNumeric(Empty)=True 같은 관용 흡수를 막기 위해 Empty/Null/배열/객체/Error를
' 명시 거부하고 숫자 판정만 IsNumeric에 맡긴다 (Date도 비숫자로 거부).
Private Sub ResolveComputeOperand(ByVal vOperand As Variant, ByRef bConst As Boolean, _
                                  ByRef idxCol As Long, ByRef vConst As Variant)
    If VarType(vOperand) = vbString Then
        If Not m_dicHeaders.exists(UCase(vOperand)) Then
            Err.Raise 9999, "clsTable.Compute", "Column not found: " & vOperand
        End If
        idxCol = m_dicHeaders(UCase(vOperand))
        bConst = False
    Else
        If IsEmpty(vOperand) Or IsNull(vOperand) Or IsArray(vOperand) _
           Or IsObject(vOperand) Or IsError(vOperand) Then
            Err.Raise 9999, "clsTable.Compute", "Operand must be a column name (String) or numeric constant"
        End If
        If Not IsNumeric(vOperand) Then
            Err.Raise 9999, "clsTable.Compute", "Operand must be a column name (String) or numeric constant"
        End If
        bConst = True
        vConst = vOperand
    End If
End Sub

Public Property Get Value(ByVal lRow As Long, ByVal sColName As String) As Variant
    Dim lCol As Long
    sColName = UCase(Trim(sColName))
    
    If m_dicHeaders.exists(sColName) Then
        lCol = m_dicHeaders(sColName)
        If lRow >= LBound(m_vData, 1) And lRow <= m_lRowCount Then
            Value = m_vData(lRow, lCol)
        End If
    End If
End Property

' [Internal] 퀵 정렬 알고리즘 (Recursive)
Private Sub QuickSort(ByRef arrIndices() As Long, ByVal L As Long, ByVal R As Long, _
                      ByVal lKeyCol As Long, ByVal bDesc As Boolean)
    Dim i As Long, j As Long
    Dim vPivot As Variant, vVal As Variant
    Dim lTemp As Long
    
    If L >= R Then Exit Sub
    
    i = L
    j = R
    
    ' 피벗 선택 (중간값)
    ' 주의: arrIndices(x)는 실제 m_vData의 행 번호를 가리킴
    vPivot = m_vData(arrIndices((L + R) \ 2), lKeyCol)
    
    Do While i <= j
        ' 오름차순/내림차순 분기
        If Not bDesc Then
            ' Ascending
            Do While CompareValues(m_vData(arrIndices(i), lKeyCol), vPivot) < 0 And i < R: i = i + 1: Loop
            Do While CompareValues(m_vData(arrIndices(j), lKeyCol), vPivot) > 0 And j > L: j = j - 1: Loop
        Else
            ' Descending
            Do While CompareValues(m_vData(arrIndices(i), lKeyCol), vPivot) > 0 And i < R: i = i + 1: Loop
            Do While CompareValues(m_vData(arrIndices(j), lKeyCol), vPivot) < 0 And j > L: j = j - 1: Loop
        End If
        
        If i <= j Then
            ' Swap Indices (데이터 이동 없음, 인덱스만 교환)
            lTemp = arrIndices(i)
            arrIndices(i) = arrIndices(j)
            arrIndices(j) = lTemp
            i = i + 1
            j = j - 1
        End If
    Loop
    
    If L < j Then Call QuickSort(arrIndices, L, j, lKeyCol, bDesc)
    If i < R Then Call QuickSort(arrIndices, i, R, lKeyCol, bDesc)
End Sub

' [Internal] 값 비교 헬퍼 (숫자 vs 문자 처리)
' 결과: -1(v1 < v2), 0(Same), 1(v1 > v2)
Private Function CompareValues(ByVal v1 As Variant, ByVal v2 As Variant) As Integer
    ' Error=safe 계약(D1): Error 셀은 raw CStr로 locale 문자열화돼 은밀히 정렬에 끼지 않도록
    ' 결정적으로 맨 뒤에 배치한다. (아래 CStr fallback은 CStr(vbError) 예외 위험도 있었다)
    Dim e1 As Boolean, e2 As Boolean
    e1 = IsError(v1): e2 = IsError(v2)
    If e1 Or e2 Then
        If e1 And e2 Then
            CompareValues = 0
        ElseIf e1 Then
            CompareValues = 1
        Else
            CompareValues = -1
        End If
        Exit Function
    End If

    ' Date=numeric 계약(D3): Date 두 값은 직렬값(CDbl) 수치 비교로 시간순을 보장한다.
    ' (CStr 텍스트 fallback은 locale 의존이라 시간순이 조용히 깨진다. 문자 혼합은 현행 텍스트 규칙 유지)
    If VarType(v1) = vbDate And VarType(v2) = vbDate Then
        If CDbl(v1) < CDbl(v2) Then
            CompareValues = -1
        ElseIf CDbl(v1) > CDbl(v2) Then
            CompareValues = 1
        Else
            CompareValues = 0
        End If
    ' If one value is a native string, force string comparison for consistency.
    ElseIf VarType(v1) = vbString Or VarType(v2) = vbString Then
        CompareValues = StrComp(CStr(v1), CStr(v2), vbTextCompare)
    ElseIf IsNumeric(v1) And IsNumeric(v2) Then
        ' Both are numeric types (but not native strings)
        If CDbl(v1) < CDbl(v2) Then
            CompareValues = -1
        ElseIf CDbl(v1) > CDbl(v2) Then
            CompareValues = 1
        Else
            CompareValues = 0
        End If
    Else
        ' Fallback for other types (e.g., numeric vs empty/error)
        CompareValues = StrComp(CStr(v1), CStr(v2), vbTextCompare)
    End If
End Function

' [Internal] 정렬된 인덱스로 데이터 재배열
Private Sub RebuildData(ByRef arrSortedIndices() As Long)
    Dim vNewData As Variant
    Dim i As Long, j As Long
    Dim lCols As Long
    
    lCols = m_lColCount
    ReDim vNewData(1 To m_lRowCount, 1 To lCols)
    
    ' 1. 헤더 복사
    For i = 1 To m_lHeaderRowIndex
        For j = 1 To lCols
            vNewData(i, j) = m_vData(i, j)
        Next j
    Next i
    
    ' 2. 데이터 복사 (정렬된 순서대로)
    For i = 0 To UBound(arrSortedIndices)
        Dim lOldRow As Long
        lOldRow = arrSortedIndices(i)
        
        For j = 1 To lCols
            ' 새 배열의 행 인덱스 = (헤더행 + 1 + i)
            vNewData(m_lHeaderRowIndex + 1 + i, j) = m_vData(lOldRow, j)
        Next j
    Next i
    
    ' 3. 교체
    m_vData = vNewData

    ' 행 순서가 바뀌었으므로 세대 증가 (전체 파생물 실효)
    BumpGeneration
End Sub
