Attribute VB_Name = "LogShim"
' @desc 디버그 로그와 테스트 어서션을 jsonl로 직접 기록 담당
Option Explicit

' vbaflow 제공 self-contained 버전. 애드인 위임(_UnivDbg.xlam) 제거 (L-H10).
' 경로: <워크북.Path>\_dbg\<워크북베이스명>.jsonl
' 레코드 스키마: kind D/J/A/X (기존 유지). 기존 호출부 시그니처 무수정.
' 기록은 진짜 append (Binary Access Read Write + LOF+1, L-H9).

Private Const DBG_FOLDER As String = "_dbg"
' 롤링 누적 상한. Reset 시에만 오래된 앞줄부터 트림한다.
Private Const MAX_LINES As Long = 5000

' 배포 무해화 스위치(설계 9.3). dist가 주입 텍스트에서만 True로 뒤집어 각 Public 싱크를
' no-op으로 만든다. 정본과 src는 항상 False(아티팩트는 빌드 산출물) - 직접 True로 두지 말 것.
' 호출부(LogShim.LogD ...)는 그대로 남아 컴파일 안전, 런타임 비용은 상수 분기 1회.
Private Const IS_RELEASE As Boolean = False

' @desc level과 msg를 D 레코드로 기록함
Public Sub LogD(level As String, msg As String)
    If IS_RELEASE Then Exit Sub
    On Error GoTo Done
    AppendLine "{""ts"":""" & NowTimestamp() & """,""kind"":""D"",""level"":""" & _
               JsonEscape(level) & """,""msg"":""" & JsonEscape(msg) & """}"
Done:
End Sub

' @desc eventName과 dataJson을 J 레코드로 기록함
Public Sub LogJ(eventName As String, dataJson As String)
    If IS_RELEASE Then Exit Sub
    On Error GoTo Done
    Dim dataPart As String
    If LenB(dataJson) = 0 Then dataPart = "null" Else dataPart = dataJson
    AppendLine "{""ts"":""" & NowTimestamp() & """,""kind"":""J"",""event"":""" & _
               JsonEscape(eventName) & """,""data"":" & dataPart & "}"
Done:
End Sub

' @desc 이벤트명과 키/값 배열을 JSON으로 조립해 LogJ로 기록함
Public Sub LogEvent(eventName As String, keys As Variant, vals As Variant)
    If IS_RELEASE Then Exit Sub
    On Error GoTo Done
    Dim json As String, i As Long
    json = "{"
    For i = LBound(keys) To UBound(keys)
        If i > LBound(keys) Then json = json & ","
        json = json & """" & JsonEscape(CStr(keys(i))) & """:" & JsonVal(vals(i))
    Next i
    json = json & "}"
    LogJ eventName, json
Done:
End Sub

' @desc 단언 결과를 A 레코드로 기록함, 실패해도 raise하지 않음
Public Sub AssertEq(name As String, expected As Variant, actual As Variant)
    If IS_RELEASE Then Exit Sub
    On Error GoTo Done
    Dim okFlag As Boolean
    okFlag = (CStr(expected) = CStr(actual))
    AppendLine "{""ts"":""" & NowTimestamp() & """,""kind"":""A"",""name"":""" & _
               JsonEscape(name) & """,""ok"":" & IIf(okFlag, "true", "false") & _
               ",""expected"":""" & JsonEscape(CStr(expected)) & _
               """,""actual"":""" & JsonEscape(CStr(actual)) & """}"
Done:
End Sub

' @desc jsonl을 상한까지 트림하고 reset 마커를 추가함
Public Sub Reset()
    If IS_RELEASE Then Exit Sub
    On Error GoTo Done
    Dim path As String
    path = LogPath()
    If LenB(path) = 0 Then Exit Sub
    EnsureDbgFolder path

    Dim marker As String
    marker = "{""ts"":""" & NowTimestamp() & """,""kind"":""X"",""event"":""reset""}"

    Dim existing As String
    If LenB(Dir(path)) > 0 Then existing = ReadFileUtf8(path)
    If Len(existing) = 0 Then
        WriteFileUtf8 path, marker & vbLf
        Exit Sub
    End If

    Dim lines() As String
    lines = Split(existing, vbLf)
    Dim keep As Long
    keep = MAX_LINES - 1
    If keep < 0 Then keep = 0

    Dim startIdx As Long, i As Long, sb As String
    Dim total As Long
    total = UBound(lines) - LBound(lines) + 1
    If total > keep Then
        startIdx = UBound(lines) - keep + 1
        If startIdx < LBound(lines) Then startIdx = LBound(lines)
    Else
        startIdx = LBound(lines)
    End If
    For i = startIdx To UBound(lines)
        If Len(lines(i)) > 0 Then sb = sb & lines(i) & vbLf
    Next i
    WriteFileUtf8 path, sb & marker & vbLf
Done:
End Sub

' === 내부 구현 ===

Private Function LogPath() As String
    ' 미저장 워크북(Path 없음)은 조용히 무시
    If LenB(ThisWorkbook.Path) = 0 Then Exit Function
    LogPath = ThisWorkbook.Path & "\" & DBG_FOLDER & "\" & BaseName(ThisWorkbook.Name) & ".jsonl"
End Function

Private Function BaseName(fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 0 Then BaseName = Left$(fileName, p - 1) Else BaseName = fileName
End Function

Private Sub EnsureDbgFolder(path As String)
    Dim folder As String
    folder = Left$(path, InStrRev(path, "\") - 1)
    If LenB(Dir(folder, vbDirectory)) = 0 Then MkDir folder
End Sub

Private Sub AppendLine(line As String)
    Dim path As String
    path = LogPath()
    If LenB(path) = 0 Then Exit Sub
    EnsureDbgFolder path

    Dim bytes() As Byte
    bytes = Utf8Bytes(line & vbLf)

    Dim f As Integer
    f = FreeFile
    On Error GoTo CleanFail
    Open path For Binary Access Read Write As #f

    Dim pos As Long
    pos = LOF(f) + 1
    ' 외부 도구가 개행 없이 끝낸 파일에 한 줄로 붙는 것 방지
    If pos > 1 Then
        Dim lastByte As Byte
        Get #f, pos - 1, lastByte
        If lastByte <> 10 Then
            Dim lfByte(0 To 0) As Byte
            lfByte(0) = 10
            Put #f, pos, lfByte
            pos = pos + 1
        End If
    End If
    Put #f, pos, bytes
    Close #f
    Exit Sub

CleanFail:
    Dim savedNum As Long, savedDesc As String
    savedNum = Err.Number: savedDesc = Err.Description
    On Error Resume Next
    Close #f
    On Error GoTo 0
    Err.Raise savedNum, "LogShim.AppendLine", savedDesc
End Sub

Private Function Utf8Bytes(s As String) As Byte()
    Dim t As Object
    Set t = CreateObject("ADODB.Stream")
    t.Type = 2
    t.Charset = "utf-8"
    t.Open
    t.WriteText s
    t.Position = 0
    t.Type = 1
    t.Position = 3 ' UTF-8 BOM 3바이트 스킵
    Utf8Bytes = t.Read
    t.Close
End Function

Private Function ReadFileUtf8(path As String) As String
    Dim s As Object
    Set s = CreateObject("ADODB.Stream")
    s.Type = 2
    s.Charset = "utf-8"
    s.Open
    s.LoadFromFile path
    Dim text As String
    text = s.ReadText
    s.Close
    If Len(text) > 0 Then
        If Left$(text, 1) = ChrW$(&HFEFF&) Then text = Mid$(text, 2)
    End If
    ReadFileUtf8 = text
End Function

Private Sub WriteFileUtf8(path As String, content As String)
    Dim t As Object, b As Object
    Set t = CreateObject("ADODB.Stream")
    t.Type = 2
    t.Charset = "utf-8"
    t.Open
    t.WriteText content

    Set b = CreateObject("ADODB.Stream")
    b.Type = 1
    b.Open
    t.Position = 3 ' UTF-8 BOM 3바이트 스킵
    t.CopyTo b
    t.Close
    b.SaveToFile path, 2 ' adSaveCreateOverWrite
    b.Close
End Sub

Private Function NowTimestamp() As String
    Dim ms As Long
    ms = CLng((Timer - Int(Timer)) * 1000) Mod 1000
    NowTimestamp = Format$(Now, "yyyy-mm-dd hh:nn:ss") & "." & Format$(ms, "000")
End Function

' @desc 값 타입별 JSON 표현 생성 (문자열만 escape, 숫자/불린은 raw)
Private Function JsonVal(v As Variant) As String
    Select Case VarType(v)
        Case vbBoolean
            JsonVal = IIf(v, "true", "false")
        Case vbInteger, vbLong, vbDouble, vbSingle, vbCurrency
            JsonVal = CStr(v)
        Case Else
            JsonVal = """" & JsonEscape(CStr(v)) & """"
    End Select
End Function

Private Function JsonEscape(s As String) As String
    Dim i As Long, ch As String, code As Long, result As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        Select Case code
            Case 34:      result = result & "\"""
            Case 92:      result = result & "\\"
            Case 8:       result = result & "\b"
            Case 9:       result = result & "\t"
            Case 10:      result = result & "\n"
            Case 12:      result = result & "\f"
            Case 13:      result = result & "\r"
            Case 0 To 31: result = result & "\u" & Right$("0000" & Hex$(code), 4)
            Case Else:    result = result & ch
        End Select
    Next i
    JsonEscape = result
End Function
