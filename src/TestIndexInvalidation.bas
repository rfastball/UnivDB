Attribute VB_Name = "TestIndexInvalidation"
Option Explicit

' Phase 1 tests for stale index prevention.

Public Sub Test_IndexInvalidation_ValueLet_RemovesNameAndIndexCaches()
    Dim T As clsTable

    Set T = NewIndexTable()

    LogShim.AssertEq "idx_valuelet_pre_name_cache", 2, T.SelectRows("Dept", "IT").Count
    LogShim.AssertEq "idx_valuelet_pre_index_cache", 1, T.SelectRowsByIndex(2, "HR").Count

    T.Value(2, "Dept") = "HR"

    LogShim.AssertEq "idx_valuelet_old_key", 1, T.SelectRows("Dept", "IT").Count
    LogShim.AssertEq "idx_valuelet_new_key", 2, T.SelectRowsByIndex(2, "HR").Count
End Sub

Public Sub Test_IndexInvalidation_UpdateBatch_RemovesNumericCache()
    Dim T As clsTable

    Set T = NewIndexTable()

    LogShim.AssertEq "idx_updatebatch_pre_index_cache", 2, T.SelectRowsByIndex(2, "IT").Count

    T.UpdateBatch "Dept", "HR", "ID", 1

    LogShim.AssertEq "idx_updatebatch_old_key", 1, T.SelectRowsByIndex(2, "IT").Count
    LogShim.AssertEq "idx_updatebatch_new_key", 2, T.SelectRows("Dept", "HR").Count
End Sub

Public Sub Test_IndexInvalidation_UpdateOnlyUpsert_ClearsCache()
    Dim tTarget As clsTable
    Dim tSource As clsTable

    Set tTarget = NewIndexTable()
    Set tSource = NewUpsertSource()

    LogShim.AssertEq "idx_upsert_pre_cache", 2, tTarget.SelectRows("Dept", "IT").Count

    tTarget.UpsertFrom tSource, "ID"

    LogShim.AssertEq "idx_upsert_old_key", 1, tTarget.SelectRows("Dept", "IT").Count
    LogShim.AssertEq "idx_upsert_new_key", 2, tTarget.SelectRowsByIndex(2, "HR").Count
End Sub

Public Sub Test_IndexInvalidation_ColumnTransforms_ClearAffectedCaches()
    Dim T As clsTable

    Set T = NewIndexTable()

    LogShim.AssertEq "idx_transform_pre_cache", 2, T.SelectRows("Dept", "IT").Count

    T.CopyColumn "Dept", "DeptCopy"
    T.Value(2, "DeptCopy") = "Changed"

    LogShim.AssertEq "idx_transform_copy_old_key", 1, T.SelectRows("DeptCopy", "IT").Count
    LogShim.AssertEq "idx_transform_copy_new_key", 1, T.SelectRowsByIndex(3, "Changed").Count

    T.Compute "Label", "Dept", "&", "DeptCopy"
    LogShim.AssertEq "idx_transform_compute_key", 1, T.SelectRows("Label", "ITIT").Count
End Sub

Private Function NewIndexTable() As clsTable
    Dim vData(1 To 4, 1 To 2) As Variant
    Dim T As clsTable

    vData(1, 1) = "ID": vData(1, 2) = "Dept"
    vData(2, 1) = 1: vData(2, 2) = "IT"
    vData(3, 1) = 2: vData(3, 2) = "IT"
    vData(4, 1) = 3: vData(4, 2) = "HR"

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData
    Set NewIndexTable = T
End Function

Private Function NewUpsertSource() As clsTable
    Dim vData(1 To 2, 1 To 2) As Variant
    Dim T As clsTable

    vData(1, 1) = "ID": vData(1, 2) = "Dept"
    vData(2, 1) = 1: vData(2, 2) = "HR"

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData
    Set NewUpsertSource = T
End Function

' 값 변경은 rowID 세대를 유지하되 보관된 인덱스의 현재성 검사는 실패한다.
Public Sub Test_IndexInvalidation_RetainedIndex_Parametric()
    Dim T As clsTable, idx As DbHashIndex, other As DbHashIndex, rows As DbRowList
    Dim backend As Long, gen As Long, n As Long
    For backend = 0 To 1
        Set T = NewIndexTable()
        T.IndexBackend = backend
        Set idx = T.GetHashIndex("Dept")
        Set other = T.GetHashIndex("ID")
        idx.TryGetRows "IT", rows
        gen = T.Generation
        T.Value(2, "Dept") = "HR"
        LogShim.AssertEq "retained_index_row_generation", gen, T.Generation
        rows.EnsureCurrent T
        other.EnsureCurrent T
        On Error Resume Next
        Err.Clear
        idx.EnsureCurrent T
        n = Err.Number
        On Error GoTo 0
        LogShim.AssertEq "retained_index_stale", 9998, n
        On Error Resume Next
        Err.Clear
        idx.TryGetRows "IT", rows
        n = Err.Number
        On Error GoTo 0
        LogShim.AssertEq "retained_index_read_stale", 9998, n
        LogShim.AssertEq "retained_index_rebuilt", 2, T.SelectRows("Dept", "HR").Count
    Next backend
End Sub
