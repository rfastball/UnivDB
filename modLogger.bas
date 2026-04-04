Attribute VB_Name = "modLogger"
Option Explicit

' === 설정 영역 ===
Private Const LOG_FILE_NAME As String = "AppLog.txt"
Private Const ENABLE_FILE_LOGGING As Boolean = True  ' 파일 저장 여부
Private Const MAX_LOG_SIZE As Long = 5242880

' [NEW] 전체 로그 시스템 ON/OFF (False면 아무것도 출력 안 함)
Private Const ENABLE_ALL_LOGGING As Boolean = False

' [NEW] 최소 로그 레벨 설정 (이 레벨 이상만 기록)
' 0: Info(전체), 1: Warning, 2: Error, 3: Debug(개발용)
' 예: 2로 설정하면 Info, Warning은 무시하고 Error만 기록
Private Const MIN_LOG_LEVEL As Integer = 0

' === 로그 레벨 열거형 ===
Public Enum LogLevel
    lvlInfo = 0     ' 일반 정보
    lvlWarning = 1  ' 경고
    lvlError = 2    ' 에러
    lvlDebug = 3    ' 디버깅
End Enum

' ========================================================
' 메인 로그 함수 (외부에서 호출하는 함수)
' ========================================================
Public Sub Log(ByVal Message As String, Optional ByVal Level As LogLevel = lvlInfo)
    ' [NEW] 1. 마스터 스위치 확인: 꺼져있으면 즉시 종료
    If Not ENABLE_ALL_LOGGING Then Exit Sub
    
    ' [NEW] 2. 로그 레벨 필터링: 설정한 레벨보다 낮으면 무시
    ' (단, Debug 레벨은 특수하게 처리하거나, 배포 시엔 lvlInfo(0)보다 높게 설정하여 숨김)
    If Level < MIN_LOG_LEVEL Then Exit Sub
    Dim strLevel As String
    Dim logOutput As String
    
    ' 1. 로그 레벨 텍스트 변환
    Select Case Level
        Case lvlInfo:    strLevel = "[INFO]   "
        Case lvlWarning: strLevel = "[WARN]   "
        Case lvlError:   strLevel = "[ERROR]  "
        Case lvlDebug:   strLevel = "[DEBUG]  "
    End Select
    
    ' 2. 출력 포맷: [시간] [레벨] 메시지
    logOutput = Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & strLevel & Message
    
    ' 3. 직접 실행 창 출력
    Debug.Print logOutput
    
    ' 4. 파일 저장 (옵션 켜져 있을 경우)
    If ENABLE_FILE_LOGGING Then
        AppendToLogFile logOutput
    End If
End Sub

' ========================================================
' 내부 함수 1: 텍스트 파일에 내용 추가
' ========================================================
Private Sub AppendToLogFile(ByVal logContent As String)
    Dim fileNum As Integer
    Dim filePath As String
    
    On Error GoTo ErrorHandler
    
    ' 파일 경로 설정
    If ThisWorkbook.Path = "" Then
        filePath = Application.DefaultFilePath & Application.PathSeparator & LOG_FILE_NAME
    Else
        filePath = ThisWorkbook.Path & Application.PathSeparator & LOG_FILE_NAME
    End If
    
    ' [추가된 기능] 파일 크기 체크 및 로테이션(백업) 수행
    CheckAndRotateLog filePath
    
    ' 파일 쓰기
    fileNum = FreeFile
    Open filePath For Append As #fileNum
    Print #fileNum, logContent
    Close #fileNum
    
    Exit Sub

ErrorHandler:
    Debug.Print "!!! 로그 파일 기록 실패 !!! : " & Err.Description
    Reset
End Sub

' ========================================================
' 내부 함수 2: 로그 파일 용량 체크 및 백업 (로테이션)
' ========================================================
Private Sub CheckAndRotateLog(ByVal filePath As String)
    Dim currentSize As Long
    Dim backupPath As String
    Dim fso As Object
    
    On Error Resume Next ' 파일이 없거나 액세스 불가시 무시
    
    ' 파일이 존재하지 않으면 체크할 필요 없음
    If Dir(filePath) = "" Then Exit Sub
    
    ' 파일 크기 확인 (FileLen 함수 사용)
    currentSize = FileLen(filePath)
    
    ' 설정한 크기(5MB)를 초과했을 경우
    If currentSize > MAX_LOG_SIZE Then
        ' 백업 파일명 생성: AppLog_20240521_143005.txt
        ' (확장자 앞에 날짜시간을 붙임)
        backupPath = Replace(filePath, ".txt", "_" & Format(Now, "yyyymmdd_hhnnss") & ".txt")
        
        ' 기존 파일의 이름을 백업 파일명으로 변경 (Rename)
        Name filePath As backupPath
        
        ' 안내 메시지를 디버그 창에 띄움
        Debug.Print "=== [System] 로그 파일 용량 초과로 백업됨: " & backupPath & " ==="
    End If
    
    On Error GoTo 0
End Sub
