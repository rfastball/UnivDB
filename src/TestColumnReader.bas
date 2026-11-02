Attribute VB_Name = "TestColumnReader"
' @desc 컬럼 reader 표면(값 패리티, stale fail-fast) 검증 담당
' @depends modDbContract, UnivDB
Option Explicit

' Phase 6.8 tests for the sealed column access surface.

Public Sub Test_ColumnReader_ValueParity()
    Dim T As clsTable
    Dim rdr As DbColumnReader
    Dim vCol As Variant
    Dim R As Long

    Set T = NewReaderTable()

    Set rdr = T.GetColumnReader("Dept")
    For R = 2 To T.RowCount
        LogShim.AssertEq "rdr_parity_dept_" & R, T.GetValue(R, "Dept"), rdr.ValueAt(R)
    Next R

    Set rdr = T.GetColumnReader("Score")
    LogShim.AssertEq "rdr_parity_num", T.GetValue(2, "Score"), rdr.ValueAt(2)
    LogShim.AssertEq "rdr_parity_empty", True, IsEmpty(rdr.ValueAt(4))

    vCol = rdr.Values
    LogShim.AssertEq "rdr_values_lbound", 2, LBound(vCol)
    LogShim.AssertEq "rdr_values_ubound", T.RowCount, UBound(vCol)
    LogShim.AssertEq "rdr_first_row", 2, rdr.FirstDataRow
    LogShim.AssertEq "rdr_last_row", T.RowCount, rdr.LastDataRow
End Sub

Public Sub Test_ColumnReader_MissingColumn_ReturnsNothing()
    Dim T As clsTable

    Set T = NewReaderTable()
    LogShim.AssertEq "rdr_missing_nothing", True, T.GetColumnReader("Nope") Is Nothing
    LogShim.AssertEq "rdr_oob_low_nothing", True, T.GetColumnReaderByIndex(0) Is Nothing
    LogShim.AssertEq "rdr_oob_high_nothing", True, T.GetColumnReaderByIndex(99) Is Nothing
End Sub

Public Sub Test_ColumnReader_StaleAfterStructuralChange_FailFast()
    Dim T As clsTable
    Dim rdr As DbColumnReader
    Dim lErr As Long

    Set T = NewReaderTable()
    Set rdr = T.GetColumnReader("Dept")

    T.AddColumn "Extra" ' 구조 변경 - 이전 reader 실효

    On Error Resume Next
    Err.Clear
    rdr.EnsureCurrent T
    lErr = Err.Number
    On Error GoTo 0

    LogShim.AssertEq "rdr_stale_err", modDbContract.DB_ERR_STALE_GENERATION, lErr

    ' 재취득한 reader는 현재 세대 - 통과해야 함
    Set rdr = T.GetColumnReader("Dept")
    On Error Resume Next
    Err.Clear
    rdr.EnsureCurrent T
    lErr = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "rdr_fresh_ok", 0, lErr
End Sub

Public Sub Test_ColumnReader_ScanSmoke()
    Dim T As clsTable
    Dim tF As clsTable
    Dim tG As clsTable
    Dim t0 As Double
    Dim lFilterMs As Long, lGroupMs As Long

    Set T = NewBigTable(5000)

    t0 = Timer
    Set tF = UnivDB.Query(T).Filter("Score", ">", 500).SelectTable()
    lFilterMs = CLng((Timer - t0) * 1000)

    t0 = Timer
    Set tG = T.GroupBy("Dept", "Score:SUM")
    lGroupMs = CLng((Timer - t0) * 1000)

    ' 회귀 감지용 기준점 기록 (Phase 7 벤치 전 임시 관측 지표)
    LogShim.LogJ "bench_seal", "{""rows"":5000,""filter_ms"":" & lFilterMs & ",""groupby_ms"":" & lGroupMs & "}"

    LogShim.AssertEq "rdr_smoke_filter_nonempty", True, tF.RowCount > 1
    LogShim.AssertEq "rdr_smoke_group_rows", 6, tG.RowCount ' 그룹 5개 + 헤더
End Sub

Private Function NewReaderTable() As clsTable
    Dim vData(1 To 4, 1 To 3) As Variant
    Dim T As New clsTable

    vData(1, 1) = "ID": vData(1, 2) = "Dept": vData(1, 3) = "Score"
    vData(2, 1) = 1: vData(2, 2) = "IT": vData(2, 3) = 10
    vData(3, 1) = 2: vData(3, 2) = "HR": vData(3, 3) = 20
    vData(4, 1) = 3: vData(4, 2) = "IT" ' Score 비움 (Empty 패리티용)

    T.LoadFromArray vData
    Set NewReaderTable = T
End Function

Private Function NewBigTable(ByVal lRows As Long) As clsTable
    Dim vData() As Variant
    Dim T As New clsTable
    Dim i As Long

    ReDim vData(1 To lRows + 1, 1 To 3)
    vData(1, 1) = "ID": vData(1, 2) = "Dept": vData(1, 3) = "Score"
    For i = 1 To lRows
        vData(i + 1, 1) = i
        vData(i + 1, 2) = "D" & (i Mod 5)
        vData(i + 1, 3) = i Mod 1000
    Next i

    T.LoadFromArray vData
    Set NewBigTable = T
End Function
