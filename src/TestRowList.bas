Attribute VB_Name = "TestRowList"
Option Explicit

Public Sub Test_RowList_Append_RowAt_ToArray_Clear()
    Dim rows As DbRowList
    Dim arr As Variant

    Set rows = New DbRowList
    rows.Append 2
    rows.Append 5
    rows.Append 9

    LogShim.AssertEq "rowlist_append_count", 3, rows.Count
    LogShim.AssertEq "rowlist_rowat_1", 2, rows.RowAt(1)
    LogShim.AssertEq "rowlist_rowat_3", 9, rows.RowAt(3)

    arr = rows.ToArray()
    LogShim.AssertEq "rowlist_array_lbound", 1, LBound(arr)
    LogShim.AssertEq "rowlist_array_ubound", 3, UBound(arr)
    LogShim.AssertEq "rowlist_array_mid", 5, arr(2)

    rows.Clear
    LogShim.AssertEq "rowlist_clear_count", 0, rows.Count
End Sub

Public Sub Test_RowList_Collection_Adapter()
    Dim source As New Collection
    Dim rows As DbRowList
    Dim roundTrip As Collection

    source.Add 3
    source.Add 7

    Set rows = New DbRowList
    rows.LoadFromCollection source

    LogShim.AssertEq "rowlist_adapter_count", 2, rows.Count
    LogShim.AssertEq "rowlist_adapter_first", 3, rows.RowAt(1)
    LogShim.AssertEq "rowlist_adapter_second", 7, rows.RowAt(2)

    Set roundTrip = rows.ToCollection()
    LogShim.AssertEq "rowlist_adapter_roundtrip_count", 2, roundTrip.Count
    LogShim.AssertEq "rowlist_adapter_roundtrip_first", 3, roundTrip(1)
    LogShim.AssertEq "rowlist_adapter_roundtrip_second", 7, roundTrip(2)
End Sub

Public Sub Test_RowList_SelectRows_Compatibility()
    Dim T As clsTable
    Dim rows As DbRowList

    Set T = NewRowListTable()
    Set rows = New DbRowList
    rows.LoadFromCollection T.SelectRows("Dept", "IT")

    LogShim.AssertEq "rowlist_selectrows_count", 2, rows.Count
    LogShim.AssertEq "rowlist_selectrows_first", 2, rows.RowAt(1)
    LogShim.AssertEq "rowlist_selectrows_second", 4, rows.RowAt(2)
End Sub

Private Function NewRowListTable() As clsTable
    Dim vData(1 To 4, 1 To 2) As Variant
    Dim T As clsTable

    vData(1, 1) = "ID": vData(1, 2) = "Dept"
    vData(2, 1) = 1: vData(2, 2) = "IT"
    vData(3, 1) = 2: vData(3, 2) = "HR"
    vData(4, 1) = 3: vData(4, 2) = "IT"

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData
    Set NewRowListTable = T
End Function
