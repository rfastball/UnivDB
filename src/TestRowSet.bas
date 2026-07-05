Attribute VB_Name = "TestRowSet"
Option Explicit

Public Sub Test_RowSet_FromAllRows_ToArray_Order()
    Dim rowSet As DbRowSet
    Dim arr As Variant

    Set rowSet = New DbRowSet
    rowSet.FromAllRows 5, 2

    LogShim.AssertEq "rowset_all_count", 4, rowSet.Count
    arr = rowSet.ToArray()
    LogShim.AssertEq "rowset_all_lbound", 0, LBound(arr)
    LogShim.AssertEq "rowset_all_first", 2, arr(0)
    LogShim.AssertEq "rowset_all_last", 5, arr(3)
End Sub

Public Sub Test_RowSet_Load_Intersect_Contains()
    Dim leftRows As DbRowList
    Dim rightRows As DbRowList
    Dim rowSet As DbRowSet
    Dim arr As Variant

    Set leftRows = New DbRowList
    leftRows.Append 5
    leftRows.Append 2
    leftRows.Append 7

    Set rightRows = New DbRowList
    rightRows.Append 7
    rightRows.Append 3
    rightRows.Append 2

    Set rowSet = New DbRowSet
    rowSet.LoadFromRowList leftRows
    rowSet.IntersectWithRowList rightRows

    LogShim.AssertEq "rowset_intersect_count", 2, rowSet.Count
    LogShim.AssertEq "rowset_contains_2", True, rowSet.ContainsRow(2)
    LogShim.AssertEq "rowset_contains_5", False, rowSet.ContainsRow(5)

    arr = rowSet.ToArray()
    LogShim.AssertEq "rowset_intersect_order_first", 2, arr(0)
    LogShim.AssertEq "rowset_intersect_order_second", 7, arr(1)
End Sub

Public Sub Test_Query_RowSet_Candidate_Order_And_Update()
    Dim T As clsTable
    Dim tResult As clsTable
    Dim lCount As Long

    Set T = NewRowSetQueryTable()

    Set tResult = UnivDB.Query(T) _
        .Where("Dept", "IT") _
        .Where("Active", "Y") _
        .SelectTable()

    LogShim.AssertEq "query_rowset_where_count", 3, tResult.RowCount
    LogShim.AssertEq "query_rowset_where_first", "Ann", tResult.Value(2, "Name")
    LogShim.AssertEq "query_rowset_where_second", "Dee", tResult.Value(3, "Name")

    lCount = UnivDB.Query(T) _
        .Where("Dept", "IT") _
        .Where("Active", "Y") _
        .SetValue("Active", "N") _
        .Update()

    LogShim.AssertEq "query_rowset_update_count", 2, lCount
    LogShim.AssertEq "query_rowset_update_remaining", 0, UnivDB.Query(T).Where("Dept", "IT").Where("Active", "Y").SelectTable().RowCount - 1
End Sub

Private Function NewRowSetQueryTable() As clsTable
    Dim vData(1 To 6, 1 To 4) As Variant
    Dim T As clsTable

    vData(1, 1) = "ID": vData(1, 2) = "Name": vData(1, 3) = "Dept": vData(1, 4) = "Active"
    vData(2, 1) = 1: vData(2, 2) = "Ann": vData(2, 3) = "IT": vData(2, 4) = "Y"
    vData(3, 1) = 2: vData(3, 2) = "Bob": vData(3, 3) = "HR": vData(3, 4) = "Y"
    vData(4, 1) = 3: vData(4, 2) = "Cy": vData(4, 3) = "IT": vData(4, 4) = "N"
    vData(5, 1) = 4: vData(5, 2) = "Dee": vData(5, 3) = "IT": vData(5, 4) = "Y"
    vData(6, 1) = 5: vData(6, 2) = "Eve": vData(6, 3) = "Sales": vData(6, 4) = "Y"

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData
    Set NewRowSetQueryTable = T
End Function
