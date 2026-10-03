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
    rowSet.IntersectWithRowList Nothing
    LogShim.AssertEq "rowset_intersect_nothing", 0, rowSet.Count
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

' UnionWith: 합집합 병합 (교차 컬럼 OR의 기반 - audit C-1)
Public Sub Test_RowSet_UnionWith()
    Dim setA As New DbRowSet
    Dim setB As New DbRowSet
    Dim setEmpty As New DbRowSet
    Dim setC As New DbRowSet
    Dim arr As Variant

    setA.AddRow 5: setA.AddRow 2: setA.AddRow 7
    setB.AddRow 7: setB.AddRow 3

    setA.UnionWith setB

    ' 겹치는 행(7)은 1회만 계수, ToArray는 rowId 오름차순
    LogShim.AssertEq "rowset_union_count", 4, setA.Count
    LogShim.AssertEq "rowset_union_has2", True, setA.ContainsRow(2)
    LogShim.AssertEq "rowset_union_has3", True, setA.ContainsRow(3)
    LogShim.AssertEq "rowset_union_has5", True, setA.ContainsRow(5)
    LogShim.AssertEq "rowset_union_has7", True, setA.ContainsRow(7)
    arr = setA.ToArray()
    LogShim.AssertEq "rowset_union_arr_first", 2, arr(0)
    LogShim.AssertEq "rowset_union_arr_last", 7, arr(3)

    ' 빈 집합과의 합집합은 무변화, 빈 집합 쪽으로의 합집합은 내용 승계
    setA.UnionWith setEmpty
    LogShim.AssertEq "rowset_union_empty_noop", 4, setA.Count
    setC.UnionWith setB
    LogShim.AssertEq "rowset_union_into_empty", 2, setC.Count
End Sub

' UnionWith 세대 계약: 서로 다른 테이블에 결속된 집합 병합은 9998
Public Sub Test_RowSet_UnionWith_BindingMismatch()
    Dim tA As clsTable, tB As clsTable
    Dim setA As New DbRowSet
    Dim setB As New DbRowSet
    Dim setU As New DbRowSet
    Dim lErrNo As Long
    Dim vData(1 To 2, 1 To 1) As Variant

    vData(1, 1) = "C"
    vData(2, 1) = 1

    Set tA = UnivDB.CreateTable(): tA.LoadFromArray vData
    Set tB = UnivDB.CreateTable(): tB.LoadFromArray vData

    setA.BindTo tA: setA.AddRow 2
    setB.BindTo tB: setB.AddRow 2

    On Error Resume Next
    setA.UnionWith setB
    lErrNo = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "rowset_union_mismatch_errno", 9998, lErrNo

    ' 미결속 상대는 계약 밖 통과 (AdoptOrVerifyBinding과 대칭)
    setU.AddRow 3
    setA.UnionWith setU
    LogShim.AssertEq "rowset_union_unbound_ok", True, setA.ContainsRow(3)
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
