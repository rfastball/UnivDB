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

    LogShim.AssertEq "baseline_groupby_rowcount", 5, tResult.RowCount
    LogShim.AssertEq "baseline_groupby_cols", 3, tResult.ColCount
    LogShim.AssertEq "baseline_groupby_it_rows", 2, tResult.SelectRows("Dept", "IT").Count
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
