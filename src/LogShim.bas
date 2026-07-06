Attribute VB_Name = "LogShim"

' @desc 본 코드용 디버그 shim, 애드인 부재 시 무해 무시 동작 노출
Option Explicit

Private Const ADDIN_FILENAME As String = "_UnivDbg.xlam"
Private Const ADDIN_PREFIX As String = "_UnivDbg.xlam!UnivDbg."

' Empty=미확인, True=가용, False=미가용 (1회 Ping 후 캐시)
Private dbgEnabled As Variant

' @desc level과 msg 받아 애드인 LogD에 위임함, 애드인 부재 시 무해 무시
Public Sub LogD(level As String, msg As String)
    If Not EnsureEnabled() Then Exit Sub
    On Error Resume Next
    Application.Run ADDIN_PREFIX & "LogDForWorkbook", ThisWorkbook.FullName, level, msg
End Sub

' @desc eventName과 dataJson 받아 애드인 LogJ에 위임함, 애드인 부재 시 무해 무시
Public Sub LogJ(eventName As String, dataJson As String)
    If Not EnsureEnabled() Then Exit Sub
    On Error Resume Next
    Application.Run ADDIN_PREFIX & "LogJForWorkbook", ThisWorkbook.FullName, eventName, dataJson
End Sub

' @desc 애드인 Reset에 위임하여 jsonl 파일을 초기화함, 애드인 부재 시 무해 무시
Public Sub Reset()
    If Not EnsureEnabled() Then Exit Sub
    On Error Resume Next
    Application.Run ADDIN_PREFIX & "ResetForWorkbook", ThisWorkbook.FullName
End Sub

' @desc 이벤트명 + 키/값 쌍들을 JSON object로 조립해 LogJ로 위임.
'       호출부가 JSON을 직접 짜지 않게 하여 깨짐을 한 곳에 가둠.
'       keys/vals는 같은 길이의 1차원 배열. vals 원소는 String/Long/Boolean 가정.
Public Sub LogEvent(eventName As String, keys As Variant, vals As Variant)
    If Not EnsureEnabled() Then Exit Sub
    On Error Resume Next
    Dim json As String, i As Long
    json = "{"
    For i = LBound(keys) To UBound(keys)
        If i > LBound(keys) Then json = json & ","
        json = json & """" & ShimEscape(CStr(keys(i))) & """:" & ShimJsonVal(vals(i))
    Next i
    json = json & "}"
    LogJ eventName, json
End Sub

' @desc 테스트 단언을 애드인 UnivDbg.AssertEq에 위임함.
Public Sub AssertEq(name As String, expected As Variant, actual As Variant)
    If Not EnsureEnabled() Then Exit Sub
    Application.Run ADDIN_PREFIX & "AssertEqForWorkbook", ThisWorkbook.FullName, name, expected, actual
End Sub

' @desc 값 타입에 따라 JSON 표현 생성 (문자열만 escape+따옴표, 숫자/불린은 raw)
Private Function ShimJsonVal(v As Variant) As String
    Select Case VarType(v)
        Case vbBoolean
            ShimJsonVal = IIf(v, "true", "false")
        Case vbInteger, vbLong, vbDouble, vbSingle, vbCurrency
            ShimJsonVal = CStr(v)
        Case Else
            ShimJsonVal = """" & ShimEscape(CStr(v)) & """"
    End Select
End Function

' @desc LogShim 자체 최소 JSON escape (본체 JsonEscape는 애드인 부재 시 못 부르므로 독립 구현)
Private Function ShimEscape(s As String) As String
    Dim i As Long, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case """": r = r & "\"""
            Case "\":  r = r & "\\"
            Case vbCr: r = r & "\r"
            Case vbLf: r = r & "\n"
            Case vbTab: r = r & "\t"
            Case Else: r = r & ch
        End Select
    Next i
    ShimEscape = r
End Function

' @desc 가용성 판정 단계: Ping→Workbooks.Open 폴백→Ping 재시도. sbroenne hidden 세션의 추가기능 자동 로드 부재 대응
Private Function EnsureEnabled() As Boolean
    If Not IsEmpty(dbgEnabled) Then
        EnsureEnabled = (dbgEnabled = True)
        Exit Function
    End If

    On Error Resume Next

    ' 1차 시도: 이미 로드된 애드인 Ping
    Err.Clear
    Application.Run ADDIN_PREFIX & "Ping"
    If Err.Number = 0 Then
        dbgEnabled = True
        EnsureEnabled = True
        Exit Function
    End If

    ' 2차 시도: UserLibraryPath에서 명시 로드 후 Ping 재시도
    Err.Clear
    Dim addinPath As String
    addinPath = Application.UserLibraryPath & ADDIN_FILENAME
    If LenB(Dir(addinPath)) > 0 Then
        Workbooks.Open fileName:=addinPath
        Err.Clear
        Application.Run ADDIN_PREFIX & "Ping"
        If Err.Number = 0 Then
            dbgEnabled = True
            EnsureEnabled = True
            Exit Function
        End If
    End If

    dbgEnabled = False
    EnsureEnabled = False
End Function


