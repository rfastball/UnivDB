Attribute VB_Name = "CodeManager"
Option Explicit

' ========================================================
' [1. 메인] IMPORT: 확장자별 헤더 복구 및 타입 지정 임포트
' ========================================================
Sub ImportVBA_Final_Version()
    Dim fso As Object, importPath As String, fileName As String, ext As String
    
    importPath = ThisWorkbook.Path & "\IMPORT\"
    If Dir(importPath, vbDirectory) = "" Then
        MsgBox "IMPORT 폴더가 없습니다.", vbExclamation: Exit Sub
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = Dir(importPath & "*.*")
    
    Do While fileName <> ""
        ext = LCase(fso.GetExtensionName(fileName))
        ' .bas, .cls, .frm 파일에 대해 처리
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            ProcessSafeImport importPath & fileName, fso.GetBaseName(fileName), ext
        End If
        fileName = Dir()
    Loop
    MsgBox "모든 모듈(클래스/폼/일반)의 타입이 정확히 복구되었습니다.", vbInformation
End Sub

' [임포트 헬퍼] 확장자별 헤더 재구성 (클래스/폼 오작동 방지)
Private Sub ProcessSafeImport(fullPath As String, modName As String, ext As String)
    Dim adoStream As Object
    Dim tempPath As String
    Dim fileContent As String
    Dim header As String
    
    tempPath = Environ("TEMP") & "\" & modName & "." & ext
    
    ' 1. UTF-8로 파일 읽기
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Type = 2: adoStream.Charset = "utf-8": adoStream.Open
    adoStream.LoadFromFile fullPath
    fileContent = adoStream.ReadText: adoStream.Close
    
    ' 2. 헤더 정보가 없는 경우 확장자에 맞춰 강제 생성 (클래스/폼 구분 핵심)
    If InStr(1, fileContent, "VERSION ", vbTextCompare) = 0 Then
        Select Case ext
            Case "cls"
                ' 클래스 모듈용 필수 헤더
                header = "VERSION 1.0 CLASS" & vbCrLf & _
                         "BEGIN" & vbCrLf & _
                         "  MultiUse = -1  'True" & vbCrLf & _
                         "END" & vbCrLf & _
                         "Attribute VB_Name = """ & modName & """" & vbCrLf & _
                         "Attribute VB_GlobalNameSpace = False" & vbCrLf & _
                         "Attribute VB_Creatable = False" & vbCrLf & _
                         "Attribute VB_PredeclaredId = False" & vbCrLf & _
                         "Attribute VB_Exposed = False" & vbCrLf
            Case "frm"
                ' 폼용 필수 헤더 (주의: .frx 파일이 같은 폴더에 있어야 디자인 복구됨)
                header = "VERSION 5.00" & vbCrLf & _
                         "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} " & modName & vbCrLf & _
                         "End" & vbCrLf & _
                         "Attribute VB_Name = """ & modName & """" & vbCrLf
            Case "bas"
                ' 일반 모듈용 헤더
                header = "Attribute VB_Name = """ & modName & """" & vbCrLf
        End Select
        fileContent = header & fileContent
    End If
    
    ' 3. ANSI(KS_C_5601)로 임시 저장
    adoStream.Charset = "ks_c_5601-1987": adoStream.Open
    adoStream.WriteText fileContent
    adoStream.SaveToFile tempPath, 2: adoStream.Close
    
    ' 4. 기존 모듈 제거 후 임포트
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(modName)
    On Error GoTo 0
    
    ' 완성된 임시 파일을 임포트 (헤더를 통해 타입을 판별함)
    ThisWorkbook.VBProject.VBComponents.Import tempPath
    
    ' 5. 임시 파일 삭제
    On Error Resume Next: Kill tempPath: On Error GoTo 0
End Sub

' ========================================================
' [2. 메인] EXPORT: 모든 정보를 유지하며 UTF-8로 저장
' (폴더명: VBA_Export_YYYYMMDD_HHMMSS 형식)
' ========================================================
Sub ExportVBA_Final_Version()
    Dim vbComp As Object, savePath As String, tempPath As String
    Dim folderName As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' [수정됨] 폴더명에 현재 날짜와 시간(초 단위까지) 포함
    folderName = "VBA_Export_" & Format(Now, "yyyymmdd_hhmmss")
    savePath = ThisWorkbook.Path & "\" & folderName & "\"
    
    ' 폴더 생성
    If Dir(savePath, vbDirectory) = "" Then MkDir savePath
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' 표준모듈(1), 클래스(2), 폼(3)만 처리
        If vbComp.Type >= 1 And vbComp.Type <= 3 Then
            Dim ext As String
            Select Case vbComp.Type
                Case 1: ext = "bas"
                Case 2: ext = "cls"
                Case 3: ext = "frm"
            End Select
            
            ' 임시 파일 경로 설정
            tempPath = Environ("TEMP") & "\temp_vba." & ext
            
            ' A. VBA 내장 Export (속성 정보 포함 추출)
            vbComp.export tempPath
            
            ' B. ANSI -> UTF-8 변환하여 저장
            ConvertAnsiToUtf8 tempPath, savePath & vbComp.Name & "." & ext
            
            ' [중요 추가] 유저폼(.frm)의 경우 .frx(바이너리) 파일도 함께 이동해야 디자인이 유지됨
            If ext = "frm" Then
                Dim frxPath As String, frxDest As String
                frxPath = Replace(tempPath, ".frm", ".frx")
                frxDest = savePath & vbComp.Name & ".frx"
                If fso.FileExists(frxPath) Then
                    fso.CopyFile frxPath, frxDest
                    fso.DeleteFile frxPath ' 임시 .frx 삭제
                End If
            End If
            
            ' 임시 파일 삭제
            If fso.FileExists(tempPath) Then fso.DeleteFile tempPath
        End If
    Next vbComp
    
    MsgBox "내보내기 완료!" & vbCrLf & "폴더: " & folderName, vbInformation
End Sub

' [내보내기 헬퍼] ANSI -> UTF-8 변환
Private Sub ConvertAnsiToUtf8(srcPath As String, destPath As String)
    Dim adoStream As Object: Set adoStream = CreateObject("ADODB.Stream")
    Dim txt As String
    
    ' 1. 원본 읽기 (ANSI)
    adoStream.Type = 2: adoStream.Charset = "ks_c_5601-1987": adoStream.Open
    adoStream.LoadFromFile srcPath
    txt = adoStream.ReadText: adoStream.Close
    
    ' 2. 변환 저장 (UTF-8)
    adoStream.Charset = "utf-8": adoStream.Open
    adoStream.WriteText txt
    adoStream.SaveToFile destPath, 2: adoStream.Close
End Sub
