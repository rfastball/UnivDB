Attribute VB_Name = "TestBaseline"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
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

Public Sub Test_Baseline_AppendFrom_PreservesRows()
    Dim T As clsTable, source As clsTable, emptyTable As New clsTable
    Dim a(1 To 2, 1 To 2) As Variant, b(1 To 5, 1 To 2) As Variant
    Dim idx As DbHashIndex, rdr As DbColumnReader, q As clsQuery
    Dim backend As Long, generation As Long, n As Long, r As Long
    a(1, 1) = "ID": a(1, 2) = "Value": a(2, 1) = 1: a(2, 2) = "old"
    b(1, 1) = "value": b(1, 2) = "id"
    b(2, 1) = CVErr(xlErrNA): b(2, 2) = 1
    b(3, 1) = DateSerial(2026, 1, 2): b(3, 2) = 1
    b(4, 2) = 2: b(5, 1) = "text": b(5, 2) = 3
    For backend = dbBackendDict To dbBackendCore
        Set T = New clsTable: T.LoadFromArray a: T.IndexBackend = backend
        Set source = New clsTable: source.LoadFromArray b
        source.Value(5, "Value") = "  exact  "
        Set idx = T.GetHashIndex("ID"): Set rdr = T.GetColumnReader("Value")
        Set q = UnivDB.Query(T).Where("ID", 1)
        generation = T.Generation
        T.AppendFrom source
        LogShim.AssertEq "append_generation_" & backend, generation + 1, T.Generation
        LogShim.AssertEq "append_rows_" & backend, 6, T.RowCount
        LogShim.AssertEq "append_old_" & backend, "old", T.Value(2, "Value")
        LogShim.AssertEq "append_duplicate_" & backend, 3, T.SelectRows("ID", 1).Count
        LogShim.AssertEq "append_error_" & backend, True, IsError(T.Value(3, "Value"))
        LogShim.AssertEq "append_date_" & backend, vbDate, VarType(T.Value(4, "Value"))
        LogShim.AssertEq "append_empty_" & backend, True, IsEmpty(T.Value(5, "Value"))
        LogShim.AssertEq "append_exact_" & backend, "  exact  ", T.Value(6, "Value")
        LogShim.AssertEq "append_source_" & backend, 5, source.RowCount
        On Error Resume Next
        Err.Clear: idx.EnsureCurrent T: n = Err.Number
        On Error GoTo 0
        LogShim.AssertEq "append_stale_index_" & backend, 9998, n
        On Error Resume Next
        Err.Clear: rdr.EnsureCurrent T: n = Err.Number
        On Error GoTo 0
        LogShim.AssertEq "append_stale_reader_" & backend, 9998, n
        On Error Resume Next
        Err.Clear: q.Filter "ID", "=", 1: n = Err.Number
        On Error GoTo 0
        LogShim.AssertEq "append_stale_query_" & backend, 9998, n
        T.AppendFrom T
        LogShim.AssertEq "append_self_rows_" & backend, 11, T.RowCount
        For r = 2 To 6
            LogShim.AssertEq "append_self_order_" & backend & "_" & r, T.Value(r, "ID"), T.Value(r + 5, "ID")
        Next r
        generation = T.Generation
        T.AppendFrom emptyTable
        LogShim.AssertEq "append_empty_noop_" & backend, generation, T.Generation
    Next backend
End Sub

Public Sub Test_Baseline_AppendFrom_ValidatesBeforeWrite()
    Dim T As clsTable, source As clsTable, missing As clsTable
    Dim a(1 To 2, 1 To 2) As Variant, b(1 To 2, 1 To 2) As Variant
    Dim headerOnly(1 To 1, 1 To 2) As Variant, raw As Variant
    Dim names As Variant, name As Variant, generation As Long, n As Long
    a(1, 1) = "ID": a(1, 2) = "Value": a(2, 1) = 1: a(2, 2) = "old"
    Set T = New clsTable: T.LoadFromArray a
    generation = T.Generation
    names = Array("Other", "ID", "")
    For Each name In names
        b(1, 1) = "ID": b(1, 2) = name: b(2, 1) = 2: b(2, 2) = "new"
        raw = b: Set source = New clsTable: source.AdoptArray raw
        On Error Resume Next
        Err.Clear: T.AppendFrom source: n = Err.Number
        On Error GoTo 0
        LogShim.AssertEq "append_bad_schema_" & CStr(name), 9999, n
        LogShim.AssertEq "append_bad_rows_" & CStr(name), 2, T.RowCount
        LogShim.AssertEq "append_bad_generation_" & CStr(name), generation, T.Generation
        LogShim.AssertEq "append_bad_value_" & CStr(name), "old", T.Value(2, "Value")
    Next name
    On Error Resume Next
    Err.Clear: T.AppendFrom missing: n = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "append_nothing", 9999, n
    headerOnly(1, 1) = "Value": headerOnly(1, 2) = "ID"
    Set source = New clsTable: source.LoadFromArray headerOnly
    T.AppendFrom source
    LogShim.AssertEq "append_header_only_noop", generation, T.Generation
    Set T = New clsTable: T.LoadFromArray a, False
    T.AppendFrom T
    LogShim.AssertEq "append_headerless_rows", 4, T.RowCount
    LogShim.AssertEq "append_headerless_first", "ID", T.Value(3, "COLUMN1")
    LogShim.AssertEq "append_headerless_last", "old", T.Value(4, "COLUMN2")
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

' vbaloop dist 스탬프(vbaloop_*)를 읽는다 - 구 vbaflow_* 이름만 읽으면 배포본이 "dev"로 보고되어
' 소비 측 최소 버전 검사가 조용히 통과한다. 속성은 테스트 동안만 메모리에 추가(ephemeral은 저장 안 함).
Public Sub Test_LibVersion_ReadsVbaloopStamp()
    Dim props As Object
    Set props = ThisWorkbook.CustomDocumentProperties
    RemoveStampProps props
    LogShim.AssertEq "libversion_unstamped_dev", "dev", UnivDB.LibVersion()

    props.Add "vbaflow_BuildVersion", False, 4, "2.8.0"
    props.Add "vbaflow_BuildRev", False, 4, "aaaaaaa"
    props.Add "vbaflow_BuiltAt", False, 4, "2026-07-01 00:00:00"
    LogShim.AssertEq "libversion_vbaflow_fallback", "2.8.0+aaaaaaa (2026-07-01 00:00:00)", UnivDB.LibVersion()

    props.Add "vbaloop_BuildVersion", False, 4, "9.9.9"
    props.Add "vbaloop_BuildRev", False, 4, "abc1234"
    props.Add "vbaloop_BuiltAt", False, 4, "2026-10-31 00:00:00"
    LogShim.AssertEq "libversion_vbaloop_wins", "9.9.9+abc1234 (2026-10-31 00:00:00)", UnivDB.LibVersion()

    RemoveStampProps props
End Sub

Private Sub RemoveStampProps(ByVal props As Object)
    Dim vName As Variant
    For Each vName In Array("vbaloop_BuildVersion", "vbaloop_BuildRev", "vbaloop_BuiltAt", _
                            "vbaflow_BuildVersion", "vbaflow_BuildRev", "vbaflow_BuiltAt")
        On Error Resume Next
        props(CStr(vName)).Delete
        On Error GoTo 0
    Next vName
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
