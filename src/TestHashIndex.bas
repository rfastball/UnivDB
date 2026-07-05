Attribute VB_Name = "TestHashIndex"
Option Explicit

Public Sub Test_DictHashIndex_KeySemantics()
    Dim idx As DictHashIndex
    Dim rows As DbRowList

    Set idx = New DictHashIndex
    idx.AddRow "IT", 2
    idx.AddRow "it", 3
    idx.AddRow Empty, 4
    idx.AddRow CVErr(xlErrNA), 5

    LogShim.AssertEq "hashindex_key_count", 2, idx.CountKeys

    LogShim.AssertEq "hashindex_try_case", True, idx.TryGetRows("IT", rows)
    LogShim.AssertEq "hashindex_case_count", 2, rows.Count
    LogShim.AssertEq "hashindex_case_first", 2, rows.RowAt(1)
    LogShim.AssertEq "hashindex_case_second", 3, rows.RowAt(2)

    LogShim.AssertEq "hashindex_try_empty", True, idx.TryGetRows("", rows)
    LogShim.AssertEq "hashindex_empty_row", 4, rows.RowAt(1)

    LogShim.AssertEq "hashindex_try_error", False, idx.TryGetRows(CVErr(xlErrNA), rows)
End Sub

Public Sub Test_TableHashIndex_SelectRows_Parity()
    Dim T As clsTable
    Dim idx As DbHashIndex
    Dim rows As DbRowList

    Set T = NewHashIndexTable()
    Set idx = T.GetHashIndex("Dept")

    LogShim.AssertEq "table_hashindex_key_count", 3, idx.CountKeys
    LogShim.AssertEq "table_hashindex_hit", True, idx.TryGetRows("IT", rows)
    LogShim.AssertEq "table_hashindex_count", T.SelectRows("Dept", "IT").Count, rows.Count
    LogShim.AssertEq "table_hashindex_first", T.SelectRows("Dept", "IT")(1), rows.RowAt(1)
    LogShim.AssertEq "table_hashindex_second", T.SelectRows("Dept", "IT")(2), rows.RowAt(2)
End Sub

Public Sub Test_HashIndex_InterfaceDispatch_Parity()
    Dim idx As DbHashIndex
    Dim rows As DbRowList
    Dim cur As DbIndexCursor
    Dim lRowSum As Long

    ' DbHashIndex 인터페이스 경유 디스패치가 구체 표면과 동치인지 확인
    Set idx = New DictHashIndex
    idx.AddRow "IT", 2
    idx.AddRow "it", 3
    idx.AddRow Empty, 4
    idx.AddRow CVErr(xlErrNA), 5

    LogShim.AssertEq "iface_key_count", 2, idx.CountKeys
    LogShim.AssertEq "iface_try_case", True, idx.TryGetRows("IT", rows)
    LogShim.AssertEq "iface_case_count", 2, rows.Count
    LogShim.AssertEq "iface_try_error", False, idx.TryGetRows(CVErr(xlErrNA), rows)

    Set cur = idx.Cursor
    Do While cur.MoveNext
        lRowSum = lRowSum + cur.Rows.Count
    Loop
    LogShim.AssertEq "iface_cursor_rowsum", 3, lRowSum

    idx.Clear
    LogShim.AssertEq "iface_clear_count", 0, idx.CountKeys
End Sub

Public Sub Test_QueryWhere_UsesHashIndex_Path()
    Dim T As clsTable
    Dim tResult As clsTable

    Set T = NewHashIndexTable()
    Set tResult = UnivDB.Query(T).Where("Dept", "IT").SelectTable()

    LogShim.AssertEq "query_hashindex_where_rowcount", 3, tResult.RowCount
    LogShim.AssertEq "query_hashindex_where_first", "Ann", tResult.Value(2, "Name")
    LogShim.AssertEq "query_hashindex_where_second", "Cy", tResult.Value(3, "Name")
End Sub

Private Function NewHashIndexTable() As clsTable
    Dim vData(1 To 5, 1 To 3) As Variant
    Dim T As clsTable

    vData(1, 1) = "ID": vData(1, 2) = "Name": vData(1, 3) = "Dept"
    vData(2, 1) = 1: vData(2, 2) = "Ann": vData(2, 3) = "IT"
    vData(3, 1) = 2: vData(3, 2) = "Bob": vData(3, 3) = "HR"
    vData(4, 1) = 3: vData(4, 2) = "Cy": vData(4, 3) = "it"
    vData(5, 1) = 4: vData(5, 2) = "Dee": vData(5, 3) = Empty

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData
    Set NewHashIndexTable = T
End Function
