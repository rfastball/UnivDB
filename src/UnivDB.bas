Attribute VB_Name = "UnivDB"
Option Explicit

' [Internal] 테이블 저장소 (Key: 테이블명, Value: clsTable 객체)
Private m_dicTables As Object

' ==============================================================================
' [0] 라이브러리 식별
' ==============================================================================

' @desc 이 라이브러리의 빌드 스탬프 반환, 배포 산출물이 아니면 "dev"
' 애드인 참조는 버전을 강제할 수단이 없다(이름/GUID 바인딩뿐). 소비 측이 구버전을
' 물고 있어도 조용히 다르게 동작할 뿐이므로, 최소한 "어느 빌드인지 물어볼 수 있게" 한다.
' 소비 측은 반드시 늦은 바인딩으로 부를 것 - 조기 바인딩하면 이 함수가 없는 구버전에서
' 컴파일 오류가 나 프로젝트 전체가 멈춘다:
'   On Error Resume Next
'   sVer = Application.Run("UnivDB.xlam!UnivDB.LibVersion")   ' Err면 구버전
' 스탬프는 vbaflow dist가 산출물에 심는 CustomDocumentProperties에서 읽는다
' (BUILD_* 상수는 합성 모듈 modBuildInfo에만 있어 src에서 직접 참조할 수 없다).
Public Function LibVersion() As String
    Dim sVer As String, sRev As String, sAt As String

    On Error Resume Next
    sVer = CStr(ThisWorkbook.CustomDocumentProperties("vbaflow_BuildVersion").Value)
    sRev = CStr(ThisWorkbook.CustomDocumentProperties("vbaflow_BuildRev").Value)
    sAt = CStr(ThisWorkbook.CustomDocumentProperties("vbaflow_BuiltAt").Value)
    On Error GoTo 0

    If Len(sVer) = 0 Then
        LibVersion = "dev"
    Else
        LibVersion = sVer & "+" & sRev & " (" & sAt & ")"
    End If
End Function

' ==============================================================================
' [1] 데이터베이스 초기화 및 테이블 등록
' ==============================================================================

' DB 엔진 초기화 (필수)
Public Sub InitDB()
    If m_dicTables Is Nothing Then
        Set m_dicTables = CreateObject("Scripting.Dictionary")
    End If
End Sub

' 시트 데이터를 이름을 지정하여 테이블로 등록
' [UnivDB.bas] 수정된 RegisterTable
Public Sub RegisterTable(ByVal sTableName As String, ByVal Source As Variant)
    Dim T As New clsTable
    
    ' 소스 타입에 따라 분기 처리
    If TypeOf Source Is Range Then
        ' Range 객체인 경우 (시트 로드)
        T.LoadFromSheet Source
    ElseIf IsArray(Source) Then
        ' 배열인 경우 (메모리 직접 로드)
        T.LoadFromArray Source
    Else
        Err.Raise 9999, , "지원하지 않는 소스 타입입니다. (Range 또는 Array 필요)"
    End If
    
    If m_dicTables Is Nothing Then InitDB
    Set m_dicTables(UCase(Trim(sTableName))) = T
End Sub

' 등록된 테이블 가져오기
Public Function GetTable(ByVal sTableName As String) As clsTable
    sTableName = UCase(Trim(sTableName))
    If m_dicTables.exists(sTableName) Then
        Set GetTable = m_dicTables(sTableName)
    Else
        Err.Raise 9999, , "등록되지 않은 테이블입니다: " & sTableName
    End If
End Function

' [Factory] 외부에서 clsTable 객체 생성 (xlam 참조 시 New 키워드 불가 대응)
Public Function CreateTable() As clsTable
    Set CreateTable = New clsTable
End Function

' ==============================================================================
' [Private Helper] Variant 입력(이름 or 객체)을 clsTable로 변환
' ==============================================================================
Private Function ResolveTable(vInput As Variant) As clsTable
    If IsObject(vInput) Then
        ' 1. 객체가 들어온 경우: clsTable인지 확인 후 반환
        If TypeOf vInput Is clsTable Then
            Set ResolveTable = vInput
        Else
            Err.Raise 9999, "UnivDB", "유효하지 않은 객체 타입입니다. clsTable만 지원합니다."
        End If
    Else
        ' 2. 문자열(이름)이 들어온 경우: 저장소에서 가져오기
        ' (GetTable 함수가 내부에 구현되어 있다고 가정)
        Set ResolveTable = GetTable(CStr(vInput))
    End If
End Function

' ==============================================================================
' [2] 초간편 쿼리 및 조인 인터페이스 (Facade)
' Update: Semi-Join 지원 및 sRightPrefix 매개변수 추가
' ==============================================================================

' 특정 테이블에 대한 쿼리 시작 (Method Chaining 지원)
Public Function Query(ByVal vTarget As Variant) As clsQuery
    Dim q As New clsQuery
    Set Query = q.From(ResolveTable(vTarget))
End Function

' 리포트 빌더 시작 (Factory Method)
Public Function Report(ByVal vTarget As Variant) As clsReportBuilder
    Dim rb As New clsReportBuilder
    Set Report = rb.Init(ResolveTable(vTarget))
End Function

' 테이블 간 조인 실행
' 입력: 테이블 이름(String) 또는 테이블 객체(clsTable) 혼용 가능
' jType: jtInner(1), jtLeft(2), jtRight(3), jtFull(4), jtLeftAnti(5), jtRightAnti(6), jtLeftSemi(7), jtRightSemi(8)
' sPrefix: 오른쪽 테이블 컬럼명 충돌 방지를 위한 접두사 (예: "R_")
' rightKey: 오른쪽 조인 키 (생략 시 leftKey와 동일)
Public Function Join(ByVal vLeft As Variant, _
                     ByVal vRight As Variant, _
                     ByVal leftKey As String, _
                     Optional ByVal jType As joinType = jtInner, _
                     Optional ByVal sPrefix As String = "", _
                     Optional ByVal rightKey As String = "") As clsTable

    Dim tLeft As clsTable
    Dim tRight As clsTable

    If rightKey = "" Then rightKey = leftKey

    Set tLeft = ResolveTable(vLeft)
    Set tRight = ResolveTable(vRight)

    Set Join = JoinTable(tLeft, tRight, leftKey, rightKey, jType, sPrefix)
End Function

' 테이블 데이터를 다른 테이블에 병합 (Upsert)
Public Sub Upsert(ByVal vTarget As Variant, ByVal vSource As Variant, ByVal sKeyCol As String)
    ResolveTable(vTarget).UpsertFrom ResolveTable(vSource), sKeyCol
End Sub

' ==============================================================================
' [3] 관리 유틸리티
' ==============================================================================
' ==============================================================================
' [4] Table Builder (Staging Area) - 팩토리 메서드
' ==============================================================================

' [Factory] 빈 빌더 생성 (NewBuilder 대체)
Public Function Builder(Optional ByVal sKeyColumn As String = "") As clsTableBuilder
    Dim tb As New clsTableBuilder
    ' 키 컬럼이 지정된 경우 설정
    If Len(sKeyColumn) > 0 Then tb.KeyColumn = sKeyColumn
    Set Builder = tb
End Function

Public Function EditTable(ByVal sTableName As String, ByVal sKeyCol As String) As clsTableBuilder
    Dim B As clsTableBuilder
    Set B = New clsTableBuilder
    
    B.KeyColumn = sKeyCol
    B.InitFromTable GetTable(sTableName)
    Set EditTable = B
End Function


' 메모리 정리
Public Sub ClearDB()
    If Not m_dicTables Is Nothing Then
        m_dicTables.RemoveAll
        Set m_dicTables = Nothing
    End If
End Sub

