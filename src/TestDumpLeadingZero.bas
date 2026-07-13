Attribute VB_Name = "TestDumpLeadingZero"
' @desc DumpToSheet의 leading-zero 텍스트 보존/숫자열 무해/옵트아웃 회귀 검증 담당
' @depends UnivDB
Option Explicit

' clsTable.DumpToSheet는 leading-zero 문자열("007")을 General 셀에 .Value2로 쓰면
' Excel이 숫자 7로 강제하는 소실점이었다. 위험 열만 @서식 선고정하는 가드를 검증한다.

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

' 옵트아웃(bPreserveLeadingZeros:=False)이면 raw 고속 경로 -> "007"이 7로 강제
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
    T.DumpToSheet ws.Range("A1"), bPreserveLeadingZeros:=False

    ' 가드 없이 General 셀에 쓰면 Excel이 숫자 7로 강제
    LogShim.AssertEq "dumplz_optout_coerced", "7", CStr(ws.Range("A2").Value)

Clean:
    KillSheet ws
End Sub
