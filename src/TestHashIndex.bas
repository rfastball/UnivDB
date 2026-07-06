Attribute VB_Name = "TestHashIndex"
' @desc 해시 인덱스 백엔드 동치(키 의미론, 커서, 쿼리 경로) 검증 담당
' @depends modIndexKey, modDbContract, UnivDB
Option Explicit

' 백엔드 파라메트릭 스위트: 동일 케이스를 DictHashIndex/CoreHashIndex 양쪽에 실행해
' DbHashIndex 계약(Empty, 대소문자, 숫자 키, Error skip, 중복 키)의 동치를 보장한다.

Public Sub Test_HashIndex_KeySemantics_Parametric()
    Test_Case_KeySemantics New DictHashIndex, "dict"
    Test_Case_KeySemantics New CoreHashIndex, "core"
End Sub

Public Sub Test_HashIndex_CursorAndClear_Parametric()
    Test_Case_CursorClear New DictHashIndex, "dict"
    Test_Case_CursorClear New CoreHashIndex, "core"
End Sub

' 케이스 본체: 대소문자 무시 / Empty / Error skip / 숫자-문자 동일 키 / 중복 키
Private Sub Test_Case_KeySemantics(ByVal idx As DbHashIndex, ByVal sTag As String)
    Dim rows As DbRowList

    idx.AddRow "IT", 2
    idx.AddRow "it", 3
    idx.AddRow Empty, 4
    idx.AddRow CVErr(xlErrNA), 5
    idx.AddRow 42, 6
    idx.AddRow "42", 7

    LogShim.AssertEq "hashindex_key_count_" & sTag, 3, idx.CountKeys

    LogShim.AssertEq "hashindex_try_case_" & sTag, True, idx.TryGetRows("IT", rows)
    LogShim.AssertEq "hashindex_case_count_" & sTag, 2, rows.Count
    LogShim.AssertEq "hashindex_case_first_" & sTag, 2, rows.RowAt(1)
    LogShim.AssertEq "hashindex_case_second_" & sTag, 3, rows.RowAt(2)

    LogShim.AssertEq "hashindex_try_empty_" & sTag, True, idx.TryGetRows("", rows)
    LogShim.AssertEq "hashindex_empty_row_" & sTag, 4, rows.RowAt(1)

    LogShim.AssertEq "hashindex_try_error_" & sTag, False, idx.TryGetRows(CVErr(xlErrNA), rows)

    ' 숫자 키는 CStr 정규화로 문자 키와 합류 (백엔드 동치 계약)
    LogShim.AssertEq "hashindex_try_numeric_" & sTag, True, idx.TryGetRows(42, rows)
    LogShim.AssertEq "hashindex_numeric_count_" & sTag, 2, rows.Count

    LogShim.AssertEq "hashindex_try_miss_" & sTag, False, idx.TryGetRows("nope", rows)
End Sub

' 케이스 본체: 커서 순회(삽입 순서)와 Clear
Private Sub Test_Case_CursorClear(ByVal idx As DbHashIndex, ByVal sTag As String)
    Dim cur As DbIndexCursor
    Dim lRowSum As Long
    Dim sFirstKey As String

    idx.AddRow "IT", 2
    idx.AddRow "it", 3
    idx.AddRow Empty, 4
    idx.AddRow "HR", 5

    Set cur = idx.Cursor
    Do While cur.MoveNext
        If sFirstKey = "" And lRowSum = 0 Then sFirstKey = CStr(cur.Key)
        lRowSum = lRowSum + cur.Rows.Count
    Loop
    LogShim.AssertEq "iface_cursor_rowsum_" & sTag, 4, lRowSum
    LogShim.AssertEq "iface_cursor_firstkey_" & sTag, "IT", sFirstKey ' 삽입 순서 보존

    idx.Clear
    LogShim.AssertEq "iface_clear_count_" & sTag, 0, idx.CountKeys

    ' Clear 후 재적재 가능해야 함
    idx.AddRow "X", 9
    LogShim.AssertEq "iface_reload_count_" & sTag, 1, idx.CountKeys
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

' 테이블 레벨 백엔드 스위치: 인덱스/쿼리/고유키/stale 계약이 core 백엔드에서도 동치
Public Sub Test_TableHashIndex_CoreBackend_Parity()
    Dim T As clsTable
    Dim idx As DbHashIndex
    Dim rows As DbRowList
    Dim tResult As clsTable
    Dim vKeys As Variant
    Dim lErr As Long

    Set T = NewHashIndexTable()
    T.IndexBackend = dbBackendCore

    Set idx = T.GetHashIndex("Dept")
    LogShim.AssertEq "core_backend_is_core", True, TypeOf idx Is CoreHashIndex
    LogShim.AssertEq "core_backend_key_count", 3, idx.CountKeys
    LogShim.AssertEq "core_backend_hit", True, idx.TryGetRows("IT", rows)
    LogShim.AssertEq "core_backend_count", 2, rows.Count

    Set tResult = UnivDB.Query(T).Where("Dept", "IT").SelectTable()
    LogShim.AssertEq "core_backend_where_rowcount", 3, tResult.RowCount

    vKeys = T.GetUniqueKeys("Dept")
    LogShim.AssertEq "core_backend_uniquekeys", 3, UBound(vKeys) - LBound(vKeys) + 1
    LogShim.AssertEq "core_backend_uniquekeys_first", "IT", CStr(vKeys(0))

    ' 세대 계약: 구조 변경 후 이전 core 인덱스는 fail-fast
    T.AddColumn "Extra"
    On Error Resume Next
    Err.Clear
    idx.EnsureCurrent T
    lErr = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "core_backend_stale_err", modDbContract.DB_ERR_STALE_GENERATION, lErr
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
