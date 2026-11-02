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
