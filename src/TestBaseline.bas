Attribute VB_Name = "TestBaseline"
Option Explicit

' Phase 0 baseline tests discovered by the VBA harness.
' These tests pin current facade behavior before internal refactoring.

Public Sub Test_Baseline_Query_Filter_Update_Delete()
    Dim T As clsTable
    Dim tResult As clsTable
    Dim lCount As Long

    Set T = NewBaseTable()

    Set tResult = UnivDB.Query(T) _
        .Where("Dept", "IT") _
        .Filter("Score", ">=", 70) _
        .SelectTable()
    LogShim.AssertEq "baseline_query_where_filter_count", 3, tResult.RowCount
    LogShim.AssertEq "baseline_query_where_filter_first", "Ann", tResult.Value(2, "Name")
    LogShim.AssertEq "baseline_query_where_filter_second", "Cy", tResult.Value(3, "Name")

    lCount = UnivDB.Query(T) _
        .Where("Name", "Bob") _
        .SetValue("Dept", "Finance") _
        .Update()
    LogShim.AssertEq "baseline_update_count", 1, lCount
    LogShim.AssertEq "baseline_update_value", "Finance", T.Value(3, "Dept")

    lCount = UnivDB.Query(T) _
        .Where("Dept", "Sales") _
        .Delete()
    LogShim.AssertEq "baseline_delete_count", 1, lCount
    LogShim.AssertEq "baseline_delete_rowcount", 5, T.RowCount
End Sub

Public Sub Test_Baseline_Select_Order_Top()
    Dim T As clsTable
    Dim tResult As clsTable

    Set T = NewBaseTable()

    Set tResult = UnivDB.Query(T) _
        .SelectColumns("Name", "Score") _
        .OrderBy("Score", True) _
        .Top(2) _
        .SelectTable()

    LogShim.AssertEq "baseline_order_top_rowcount", 3, tResult.RowCount
    LogShim.AssertEq "baseline_order_top_first", "Cy", tResult.Value(2, "Name")
    LogShim.AssertEq "baseline_order_top_second", "Ann", tResult.Value(3, "Name")
End Sub

Public Sub Test_Baseline_Join_Types()
    Dim tLeft As clsTable
    Dim tRight As clsTable
    Dim tResult As clsTable

    Set tLeft = NewJoinLeftTable()
    Set tRight = NewJoinRightTable()

    Set tResult = UnivDB.Join(tLeft, tRight, "Dept", jtInner)
    LogShim.AssertEq "baseline_join_inner_rowcount", 3, tResult.RowCount

    Set tResult = UnivDB.Join(tLeft, tRight, "Dept", jtLeft)
    LogShim.AssertEq "baseline_join_left_rowcount", 5, tResult.RowCount

    Set tResult = UnivDB.Join(tLeft, tRight, "Dept", jtLeftAnti)
    LogShim.AssertEq "baseline_join_left_anti_rowcount", 3, tResult.RowCount

    Set tResult = UnivDB.Join(tLeft, tRight, "Dept", jtLeftSemi)
    LogShim.AssertEq "baseline_join_left_semi_rowcount", 3, tResult.RowCount
End Sub

Public Sub Test_Baseline_GroupBy()
    Dim T As clsTable
    Dim tResult As clsTable

    Set T = NewBaseTable()
    Set tResult = T.GroupBy("Dept", "Score:SUM,ID:COUNT")

    ' R1 결정 B: GroupBy 키도 인덱스 계약(대소문자 무시)을 따름 - "IT"/"it"는 한 그룹
    LogShim.AssertEq "baseline_groupby_rowcount", 4, tResult.RowCount
    LogShim.AssertEq "baseline_groupby_cols", 3, tResult.ColCount
    LogShim.AssertEq "baseline_groupby_it_rows", 1, tResult.SelectRows("Dept", "IT").Count
End Sub

Private Function NewBaseTable() As clsTable
    Dim vData(1 To 6, 1 To 4) As Variant
    Dim T As clsTable

    vData(1, 1) = "ID": vData(1, 2) = "Name": vData(1, 3) = "Dept": vData(1, 4) = "Score"
    vData(2, 1) = 1: vData(2, 2) = "Ann": vData(2, 3) = "IT": vData(2, 4) = 80
    vData(3, 1) = 2: vData(3, 2) = "Bob": vData(3, 3) = "HR": vData(3, 4) = 60
    vData(4, 1) = 3: vData(4, 2) = "Cy": vData(4, 3) = "IT": vData(4, 4) = 90
    vData(5, 1) = 4: vData(5, 2) = "Dee": vData(5, 3) = "Sales": vData(5, 4) = 75
    vData(6, 1) = 5: vData(6, 2) = "Eve": vData(6, 3) = "it": vData(6, 4) = 50

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData
    Set NewBaseTable = T
End Function

Private Function NewJoinLeftTable() As clsTable
    Dim vData(1 To 5, 1 To 2) As Variant
    Dim T As clsTable

    vData(1, 1) = "ID": vData(1, 2) = "Dept"
    vData(2, 1) = 1: vData(2, 2) = "A"
    vData(3, 1) = 2: vData(3, 2) = "A"
    vData(4, 1) = 3: vData(4, 2) = "B"
    vData(5, 1) = 4: vData(5, 2) = "X"

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData
    Set NewJoinLeftTable = T
End Function

Private Function NewJoinRightTable() As clsTable
    Dim vData(1 To 3, 1 To 2) As Variant
    Dim T As clsTable

    vData(1, 1) = "Dept": vData(1, 2) = "Label"
    vData(2, 1) = "A": vData(2, 2) = "Alpha"
    vData(3, 1) = "C": vData(3, 2) = "Gamma"

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData
    Set NewJoinRightTable = T
End Function

' ==============================================================================
' 로드 경계 fail-fast (LoadFromArray) - 빈/1D/스칼라 입력 + 중복 헤더 거부
' ==============================================================================

' 중복 헤더명은 조용히 덮어쓰지 않고 즉시 Err.Raise 9999 (조용히 틀린 결과 방지)
Public Sub Test_Load_DuplicateHeader_Raises()
    Dim T As clsTable
    Dim lErrNo As Long
    Dim vData(1 To 2, 1 To 3) As Variant

    vData(1, 1) = "ID": vData(1, 2) = "Name": vData(1, 3) = "ID"
    vData(2, 1) = 1: vData(2, 2) = "Ann": vData(2, 3) = 2

    Set T = UnivDB.CreateTable()
    On Error Resume Next
    T.LoadFromArray vData
    lErrNo = Err.Number
    On Error GoTo 0

    LogShim.AssertEq "load_dup_header_errno", 9999, lErrNo
End Sub

' 헤더 중복 판정은 UCase 기준(대소문자 무시) - "Dept" vs "dept"도 충돌
Public Sub Test_Load_DuplicateHeader_CaseInsensitive_Raises()
    Dim T As clsTable
    Dim lErrNo As Long
    Dim vData(1 To 2, 1 To 2) As Variant

    vData(1, 1) = "Dept": vData(1, 2) = "dept"
    vData(2, 1) = "IT": vData(2, 2) = "HR"

    Set T = UnivDB.CreateTable()
    On Error Resume Next
    T.LoadFromArray vData
    lErrNo = Err.Number
    On Error GoTo 0

    LogShim.AssertEq "load_dup_header_ci_errno", 9999, lErrNo
End Sub

' 1차원 배열은 난해한 오류 9 대신 명확한 9999로 거부
Public Sub Test_Load_1DArray_Raises()
    Dim T As clsTable
    Dim lErrNo As Long
    Dim vData(1 To 3) As Variant

    vData(1) = "A": vData(2) = "B": vData(3) = "C"

    Set T = UnivDB.CreateTable()
    On Error Resume Next
    T.LoadFromArray vData
    lErrNo = Err.Number
    On Error GoTo 0

    LogShim.AssertEq "load_1d_errno", 9999, lErrNo
End Sub

' 스칼라(비배열) 입력은 9999로 거부 (단일 셀 시트가 스칼라 Value2를 흘리는 경로 포함)
Public Sub Test_Load_ScalarOrEmpty_Raises()
    Dim T As clsTable
    Dim lErrNo As Long

    Set T = UnivDB.CreateTable()
    On Error Resume Next
    T.LoadFromArray 42
    lErrNo = Err.Number
    On Error GoTo 0

    LogShim.AssertEq "load_scalar_errno", 9999, lErrNo
End Sub

' 가드가 Clear 앞이라 잘못된 입력 로드 시도는 기존 테이블 상태를 파괴하지 않는다
Public Sub Test_Load_FailPreservesState()
    Dim T As clsTable
    Dim vGood(1 To 3, 1 To 2) As Variant
    Dim vBad(1 To 3) As Variant

    vGood(1, 1) = "ID": vGood(1, 2) = "Name"
    vGood(2, 1) = 1: vGood(2, 2) = "Ann"
    vGood(3, 1) = 2: vGood(3, 2) = "Bob"

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vGood
    LogShim.AssertEq "load_preserve_precount", 3, T.RowCount

    vBad(1) = "x": vBad(2) = "y": vBad(3) = "z"
    On Error Resume Next
    T.LoadFromArray vBad
    On Error GoTo 0

    ' 예외 후에도 기존 데이터/헤더가 그대로 보존됨
    LogShim.AssertEq "load_preserve_postcount", 3, T.RowCount
    LogShim.AssertEq "load_preserve_header", True, T.Headers.exists("NAME")

    Dim duplicate(1 To 1, 1 To 2) As Variant, gen As Long, n As Long
    duplicate(1, 1) = " Dept ": duplicate(1, 2) = "DEPT"
    gen = T.Generation
    On Error Resume Next
    Err.Clear
    T.LoadFromArray duplicate
    n = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "load_duplicate_preflight", 9999, n
    LogShim.AssertEq "load_duplicate_preserve_generation", gen, T.Generation
    LogShim.AssertEq "load_duplicate_preserve_rows", 3, T.RowCount
    LogShim.AssertEq "load_duplicate_preserve_value", "Ann", T.Value(2, "Name")
End Sub

' 내부 신뢰 경로(AdoptArray)는 중복 헤더를 거부하지 않고 last-wins 유지 (비대칭 회귀 고정)
Public Sub Test_Load_AdoptArray_DuplicateHeader_LastWins()
    Dim T As clsTable
    Dim vData(1 To 2, 1 To 2) As Variant

    vData(1, 1) = "K": vData(1, 2) = "K"
    vData(2, 1) = 10: vData(2, 2) = 20

    Set T = UnivDB.CreateTable()
    T.AdoptArray vData

    ' 거부 없이 로드되고, 동명 헤더는 마지막 열(2)이 이긴다
    LogShim.AssertEq "adopt_dup_lastwins_col", 2, T.Headers("K")
End Sub

' --- 라이브러리 식별: 소비 측이 버전을 물어볼 수 있어야 한다(애드인은 버전 강제 불가) ---
Public Sub Test_LibVersion_Reports_Something()
    Dim sVer As String
    sVer = UnivDB.LibVersion()
    LogShim.AssertEq "libversion_nonempty", True, Len(sVer) > 0
End Sub

' Headers 사본을 수정해도 원본 스키마/인덱스/값은 영향을 받지 않는다.
Public Sub Test_Headers_SnapshotIsolation()
    Dim T As clsTable, headers As Object, result As clsTable, gen As Long, n As Long, col As Long
    Set T = NewBaseTable()
    Set headers = T.Headers
    gen = T.Generation
    headers("NAME") = 999
    headers.RemoveAll
    T.Headers("NAME") = 999
    T.Headers.RemoveAll
    LogShim.AssertEq "headers_original_count", 4, T.Headers.Count
    LogShim.AssertEq "headers_original_name_col", 2, T.ColumnIndex(" name ")
    LogShim.AssertEq "headers_has_column_normalized", True, T.HasColumn(" name ")
    LogShim.AssertEq "headers_missing_column", False, T.HasColumn("__bad__")
    On Error Resume Next
    Err.Clear
    col = T.ColumnIndex("__bad__")
    n = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "headers_bad_index_failfast", 9999, n
    LogShim.AssertEq "headers_original_generation", gen, T.Generation
    Set result = UnivDB.Query(T).Where("Name", "Ann").ToTable()
    LogShim.AssertEq "headers_query_unaffected", "Ann", result.Value(2, "Name")

    Set headers = T.Headers
    T.AddColumn "Extra"
    LogShim.AssertEq "headers_snapshot_no_new_column", False, headers.exists("EXTRA")
End Sub
