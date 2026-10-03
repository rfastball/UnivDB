Attribute VB_Name = "TestDumpLeadingZero"
' @desc DumpToSheet/BuildToSheet 텍스트 가드(leading-zero/긴코드) + 기본 선클리어 계약 회귀 검증 담당
' @depends UnivDB
Option Explicit

' clsTable.DumpToSheet는 숫자꼴 문자열("007", 긴 코드)을 General 셀에 .Value2로 쓰면
' Excel이 숫자로 강제하는 소실점이었다. 위험 열만 @서식 선고정하는 가드를 검증하고,
' 리포트 경로(BuildToSheet)는 강한 가드(dbGuardAllNumericText)로 서브타입을 보존하되
' 진짜 숫자 열은 숫자로 남겨 DefineColumn 서식이 적용되는지 검증한다.

Private Function NewTempSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ' 갓 추가된 시트는 전 셀 General - 소실 조건 재현에 적합
    Set NewTempSheet = ws
End Function

Private Sub KillSheet(ByVal ws As Worksheet)
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

' 텍스트 ID 열은 보존, 같은 덤프의 숫자 열은 숫자로 유지
Public Sub Test_DumpLeadingZero_Preserved()
    Dim T As clsTable
    Dim ws As Worksheet
    Dim vData(1 To 3, 1 To 2) As Variant

    vData(1, 1) = "ID": vData(1, 2) = "Amt"
    vData(2, 1) = "007": vData(2, 2) = 12345
    vData(3, 1) = "012": vData(3, 2) = 67890

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    Set ws = NewTempSheet()
    On Error GoTo Clean
    T.DumpToSheet ws.Range("A1")

    ' ID 열: leading-zero 문자열 보존 (숫자 강제 안 됨)
    LogShim.AssertEq "dumplz_id_row2", "007", CStr(ws.Range("A2").Value)
    LogShim.AssertEq "dumplz_id_row3", "012", CStr(ws.Range("A3").Value)
    LogShim.AssertEq "dumplz_id_fmt", "@", ws.Range("A2").NumberFormat

    ' Amt 열: 위험 문자열 없음 -> @서식 강제 안 됨, 숫자로 유지 (vbDouble=5)
    LogShim.AssertEq "dumplz_amt_isnum", 5, VarType(ws.Range("B2").Value)
    LogShim.AssertEq "dumplz_amt_val", 12345, ws.Range("B2").Value
    LogShim.AssertEq "dumplz_amt_fmt_notext", "False", CStr(ws.Range("B2").NumberFormat = "@")

Clean:
    KillSheet ws
End Sub

' 옵트아웃(dbGuardNone)이면 raw 고속 경로 -> "007"이 7로 강제
Public Sub Test_DumpLeadingZero_OptOut()
    Dim T As clsTable
    Dim ws As Worksheet
    Dim vData(1 To 2, 1 To 1) As Variant

    vData(1, 1) = "ID"
    vData(2, 1) = "007"

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    Set ws = NewTempSheet()
    On Error GoTo Clean
    T.DumpToSheet ws.Range("A1"), eTextGuard:=dbGuardNone

    ' 가드 없이 General 셀에 쓰면 Excel이 숫자 7로 강제
    LogShim.AssertEq "dumplz_optout_coerced", "7", CStr(ws.Range("A2").Value)

Clean:
    KillSheet ws
End Sub

' 강한 가드(dbGuardAllNumericText): leading-zero 없는 긴 숫자 코드도 텍스트로 보존
' (leading-zero 전용 가드였다면 General로 가서 정밀도 손실됐을 케이스)
Public Sub Test_DumpGuard_AllNumericText_LongCode()
    Dim T As clsTable
    Dim ws As Worksheet
    Dim vData(1 To 2, 1 To 1) As Variant

    vData(1, 1) = "Acct"
    vData(2, 1) = "1234567890123456" ' 16자리, leading-zero 없음

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    Set ws = NewTempSheet()
    On Error GoTo Clean
    T.DumpToSheet ws.Range("A1"), eTextGuard:=dbGuardAllNumericText

    ' 텍스트로 보존되어 16자리 원문 그대로 (숫자였다면 15자리 정밀도로 truncate)
    LogShim.AssertEq "dumpguard_longcode_preserved", "1234567890123456", CStr(ws.Range("A2").Value)
    LogShim.AssertEq "dumpguard_longcode_fmt", "@", ws.Range("A2").NumberFormat

Clean:
    KillSheet ws
End Sub

' BuildToSheet 리포트 경로: 텍스트 열 보존 + 숫자 열은 진짜 숫자로 유지되어 DefineColumn 서식 적용
Public Sub Test_BuildToSheet_PreservesTextAndNumbers()
    Dim T As clsTable, loaded As New clsTable, typed As New clsTable
    Dim ws As Worksheet
    Dim vData(1 To 3, 1 To 4) As Variant, n As Long

    vData(1, 1) = "ID": vData(1, 2) = "Acct": vData(1, 3) = "Amt"
    vData(2, 1) = "007": vData(2, 2) = "1234567890123456": vData(2, 3) = 12345
    vData(3, 1) = "012": vData(3, 2) = "9876543210987654": vData(3, 3) = 67890

    vData(1, 4) = "Day"
    vData(2, 4) = DateSerial(2026, 10, 3)
    vData(3, 4) = DateSerial(2026, 10, 4)
    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    Set ws = NewTempSheet()
    On Error GoTo Clean
    UnivDB.Report(T).DefineColumn("Amt", , "#,##0").DefineColumn("Day", , "yyyy-mm-dd").BuildToSheet ws, "A1"

    ' ID: leading-zero 텍스트 보존
    LogShim.AssertEq "b2s_id_preserved", "007", CStr(ws.Range("A2").Value)
    ' Acct: 긴 숫자 코드 텍스트 보존 (정밀도 손실 없음)
    LogShim.AssertEq "b2s_acct_preserved", "1234567890123456", CStr(ws.Range("B2").Value)
    ' Amt: 진짜 숫자로 유지 (vbDouble=5) -> DefineColumn 숫자서식이 실제로 적용됨
    LogShim.AssertEq "b2s_amt_isnum", 5, VarType(ws.Range("C2").Value)
    LogShim.AssertEq "b2s_amt_val", 12345, ws.Range("C2").Value

    ' LoadFromSheet uses Value2: date format does not turn a Double into a Date.
    loaded.LoadFromSheet ws.Range("A1:D3")
    LogShim.AssertEq "sheet_date_value2_type", vbDouble, VarType(loaded.Value(2, "Day"))
    LogShim.AssertEq "sheet_date_serial_lookup", 1, loaded.SelectRows("Day", ws.Range("D2").Value2).Count
    LogShim.AssertEq "sheet_date_not_iso_lookup", 0, loaded.SelectRows("Day", "2026-10-03").Count

    ' Explicit Range.Value input preserves Excel's Date subtype.
    typed.LoadFromArray ws.Range("A1:D3").Value
    LogShim.AssertEq "sheet_date_value_type", vbDate, VarType(typed.Value(2, "Day"))
    LogShim.AssertEq "sheet_date_iso_lookup", 1, typed.SelectRows("Day", "2026-10-03").Count

Clean:
    n = Err.Number
    KillSheet ws
    LogShim.AssertEq "sheet_roundtrip_no_error", 0, n
End Sub

' 축소 재덤프: 큰 덤프 뒤 작은 덤프가 기본 선클리어로 옛 초과행을 제거해 유령행이 안 남는다.
' (하위가 CurrentRegion으로 읽어도 유령행을 흡수하지 않음)
Public Sub Test_DumpToSheet_ShrinkClearsGhostRows()
    Dim tBig As clsTable, tSmall As clsTable
    Dim ws As Worksheet
    Dim vBig(1 To 5, 1 To 1) As Variant
    Dim vSmall(1 To 2, 1 To 1) As Variant

    vBig(1, 1) = "ID"
    vBig(2, 1) = "b1": vBig(3, 1) = "b2": vBig(4, 1) = "b3": vBig(5, 1) = "b4"
    vSmall(1, 1) = "ID"
    vSmall(2, 1) = "s1"

    Set tBig = UnivDB.CreateTable(): tBig.LoadFromArray vBig
    Set tSmall = UnivDB.CreateTable(): tSmall.LoadFromArray vSmall

    Set ws = NewTempSheet()
    On Error GoTo Clean
    tBig.DumpToSheet ws.Range("A1")      ' A1..A5 채움
    tSmall.DumpToSheet ws.Range("A1")    ' 기본 선클리어 -> A3..A5 유령행 제거

    ' 옛 초과행(A4,A5) 비워짐
    LogShim.AssertEq "shrink_ghost_a4", "", CStr(ws.Range("A4").Value)
    LogShim.AssertEq "shrink_ghost_a5", "", CStr(ws.Range("A5").Value)
    ' CurrentRegion이 새 크기(헤더+1행=2)만 흡수
    LogShim.AssertEq "shrink_region_rows", 2, ws.Range("A1").CurrentRegion.Rows.Count

Clean:
    KillSheet ws
End Sub

' 옵트아웃(bClearRegion:=False): 선클리어 안 함 -> 외과적 쓰기 동작 보존(옛 초과행 잔존)
Public Sub Test_DumpToSheet_ClearOptOut_PreservesGhost()
    Dim tBig As clsTable, tSmall As clsTable
    Dim ws As Worksheet
    Dim vBig(1 To 5, 1 To 1) As Variant
    Dim vSmall(1 To 2, 1 To 1) As Variant

    vBig(1, 1) = "ID"
    vBig(2, 1) = "b1": vBig(3, 1) = "b2": vBig(4, 1) = "b3": vBig(5, 1) = "b4"
    vSmall(1, 1) = "ID"
    vSmall(2, 1) = "s1"

    Set tBig = UnivDB.CreateTable(): tBig.LoadFromArray vBig
    Set tSmall = UnivDB.CreateTable(): tSmall.LoadFromArray vSmall

    Set ws = NewTempSheet()
    On Error GoTo Clean
    tBig.DumpToSheet ws.Range("A1")
    tSmall.DumpToSheet ws.Range("A1"), bClearRegion:=False   ' 선클리어 억제

    ' 새 데이터는 사각형만 덮어씀 (A2 갱신)
    LogShim.AssertEq "optout_new_a2", "s1", CStr(ws.Range("A2").Value)
    ' 옛 초과행은 잔존 (외과적 쓰기 계약 보존)
    LogShim.AssertEq "optout_ghost_a5", "b4", CStr(ws.Range("A5").Value)

Clean:
    KillSheet ws
End Sub

' 빈 결과 덤프도 옛 영역을 비운다 (선클리어가 m_lRowCount=0 조기탈출보다 앞섬)
Public Sub Test_DumpToSheet_EmptyClearsRegion()
    Dim tSeed As clsTable, tEmpty As clsTable
    Dim ws As Worksheet
    Dim vSeed(1 To 3, 1 To 1) As Variant

    vSeed(1, 1) = "ID"
    vSeed(2, 1) = "e1": vSeed(3, 1) = "e2"

    Set tSeed = UnivDB.CreateTable(): tSeed.LoadFromArray vSeed
    Set tEmpty = UnivDB.CreateTable()   ' 0행 (m_lRowCount=0)

    Set ws = NewTempSheet()
    On Error GoTo Clean
    tSeed.DumpToSheet ws.Range("A1")    ' A1..A3 채움
    tEmpty.DumpToSheet ws.Range("A1")   ' 빈 덤프 -> 선클리어 후 조기탈출

    LogShim.AssertEq "empty_clear_a2", "", CStr(ws.Range("A2").Value)
    LogShim.AssertEq "empty_clear_a3", "", CStr(ws.Range("A3").Value)

Clean:
    KillSheet ws
End Sub
