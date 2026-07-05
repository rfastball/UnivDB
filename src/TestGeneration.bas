Attribute VB_Name = "TestGeneration"
' @desc 세대 계약(구조 변경 bump, stale 파생 구조 fail-fast) 검증 담당
' @depends modDbContract, UnivDB
Option Explicit

' Phase 6.3 tests for the table generation contract.

Public Sub Test_Generation_StructuralChanges_Bump()
    Dim T As clsTable
    Dim g As Long

    Set T = NewGenTable()
    LogShim.AssertEq "gen_load_bumped", True, T.Generation > 0

    g = T.Generation
    T.OrderBy "ID"
    LogShim.AssertEq "gen_orderby", g + 1, T.Generation

    g = T.Generation
    T.AddColumn "Extra"
    LogShim.AssertEq "gen_addcolumn", g + 1, T.Generation

    g = T.Generation
    T.RemoveColumn "Extra"
    LogShim.AssertEq "gen_removecolumn", g + 1, T.Generation

    g = T.Generation
    T.DeleteBatch "Dept", "HR"
    LogShim.AssertEq "gen_deletebatch", g + 1, T.Generation

    g = T.Generation
    T.SelectColumns "ID", "Dept"
    LogShim.AssertEq "gen_selectcolumns", g + 1, T.Generation
End Sub

Public Sub Test_Generation_Upsert_Bumps()
    Dim T As clsTable
    Dim g As Long

    Set T = NewGenTable()

    ' update-only upsert (rowID 안정적이지만 문서상 구조 변경으로 열거)
    g = T.Generation
    T.UpsertFrom NewGenSource(1, "Sales"), "ID"
    LogShim.AssertEq "gen_upsert_update_only", g + 1, T.Generation

    ' insert upsert
    g = T.Generation
    T.UpsertFrom NewGenSource(99, "Lab"), "ID"
    LogShim.AssertEq "gen_upsert_insert", g + 1, T.Generation
    LogShim.AssertEq "gen_upsert_rowcount", 5, T.RowCount
End Sub

Public Sub Test_Generation_ValueChanges_NoBump()
    Dim T As clsTable
    Dim g As Long

    Set T = NewGenTable()
    g = T.Generation

    T.Value(2, "Dept") = "HR"
    LogShim.AssertEq "gen_valuelet_nobump", g, T.Generation

    T.UpdateBatch "Dept", "Sales", "ID", 3
    LogShim.AssertEq "gen_updatebatch_nobump", g, T.Generation
End Sub

Public Sub Test_Generation_StaleRowSet_FailFast()
    Dim T As clsTable
    Dim q As clsQuery
    Dim lErr As Long

    Set T = NewGenTable()
    Set q = UnivDB.Query(T).Where("Dept", "IT")

    T.OrderBy "ID" ' 구조 변경 - 이전 후보군 실효

    On Error Resume Next
    Err.Clear
    q.SelectTable
    lErr = Err.Number
    On Error GoTo 0

    LogShim.AssertEq "gen_stale_rowset_err", modDbContract.DB_ERR_STALE_GENERATION, lErr
End Sub

Public Sub Test_Generation_StaleIndex_FailFast()
    Dim T As clsTable
    Dim idx As DictHashIndex
    Dim lErr As Long

    Set T = NewGenTable()
    Set idx = T.GetHashIndex("Dept")

    T.OrderBy "ID" ' 구조 변경 - 이전 인덱스 실효

    On Error Resume Next
    Err.Clear
    idx.EnsureCurrent T
    lErr = Err.Number
    On Error GoTo 0

    LogShim.AssertEq "gen_stale_index_err", modDbContract.DB_ERR_STALE_GENERATION, lErr

    ' 재취득한 인덱스는 현재 세대 - 통과해야 함
    Set idx = T.GetHashIndex("Dept")
    On Error Resume Next
    Err.Clear
    idx.EnsureCurrent T
    lErr = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "gen_fresh_index_ok", 0, lErr
End Sub

Public Sub Test_Generation_MixedBinding_FailFast()
    Dim T As clsTable
    Dim idx As DictHashIndex
    Dim rows As DbRowList
    Dim rowSet As DbRowSet
    Dim lErr As Long

    Set T = NewGenTable()

    ' 이전 세대의 RowList로 RowSet 구성
    Set idx = T.GetHashIndex("Dept")
    idx.TryGetRows "IT", rows
    Set rowSet = New DbRowSet
    rowSet.LoadFromRowList rows ' 결속 승계

    T.OrderBy "ID" ' 구조 변경

    ' 새 세대의 RowList와 교집합 시도 - 결속 불일치로 즉시 실패
    Set idx = T.GetHashIndex("Dept")
    idx.TryGetRows "IT", rows

    On Error Resume Next
    Err.Clear
    rowSet.IntersectWithRowList rows
    lErr = Err.Number
    On Error GoTo 0

    LogShim.AssertEq "gen_mixed_binding_err", modDbContract.DB_ERR_STALE_GENERATION, lErr
End Sub

Public Sub Test_Generation_NormalQueryPath_Unaffected()
    Dim T As clsTable
    Dim tResult As clsTable

    Set T = NewGenTable()

    Set tResult = UnivDB.Query(T) _
        .Where("Dept", "IT") _
        .Filter("Score", ">", 5) _
        .SelectTable()

    ' RowCount는 헤더 행 포함 (데이터 2행 + 헤더 1행)
    LogShim.AssertEq "gen_normal_query_count", 3, tResult.RowCount
End Sub

Private Function NewGenTable() As clsTable
    Dim vData(1 To 4, 1 To 3) As Variant
    Dim T As New clsTable

    vData(1, 1) = "ID": vData(1, 2) = "Dept": vData(1, 3) = "Score"
    vData(2, 1) = 1: vData(2, 2) = "IT": vData(2, 3) = 10
    vData(3, 1) = 2: vData(3, 2) = "HR": vData(3, 3) = 20
    vData(4, 1) = 3: vData(4, 2) = "IT": vData(4, 3) = 30

    T.LoadFromArray vData
    Set NewGenTable = T
End Function

Private Function NewGenSource(ByVal lId As Long, ByVal sDept As String) As clsTable
    Dim vData(1 To 2, 1 To 2) As Variant
    Dim T As New clsTable

    vData(1, 1) = "ID": vData(1, 2) = "Dept"
    vData(2, 1) = lId: vData(2, 2) = sDept

    T.LoadFromArray vData
    Set NewGenSource = T
End Function
