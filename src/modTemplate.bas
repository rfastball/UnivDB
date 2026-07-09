Attribute VB_Name = "modTemplate"
Option Explicit

' ==============================================================================
' [UnivDB Code Templates v2.0]
' 설명: UnivDB 라이브러리의 표준 사용 패턴 (Best Practices)
' 주요 특징:
' 1. 파사드(UnivDB.Query)를 통한 직관적인 메서드 체이닝
' 2. ReportBuilder를 활용한 데이터 가공 파이프라인 구축
' 3. 최신 CRUD(Insert, Update, Delete) 기능 활용
' ==============================================================================

' ------------------------------------------------------------------------------
' [Template 1] 기본 조회 및 추출 (Basic Query)
' 설명: UnivDB.Query 파사드를 사용하여 데이터를 조회, 필터링, 정렬 후 출력합니다.
' ------------------------------------------------------------------------------
Sub Template_01_BasicQuery()
    ' ================= [Configuration] =================
    Const T_SOURCE As String = "EmpData"
    Const T_TARGET_SHEET As String = "Report"
    Const T_TARGET_CELL As String = "A1"
    ' ===================================================

    ' 1. DB 초기화 및 데이터 등록
    UnivDB.InitDB
    UnivDB.RegisterTable T_SOURCE, Sheet1.Range("A1").CurrentRegion
    
    ' 2. 쿼리 실행 (Method Chaining)
    '    Where(인덱스) -> Filter(스캔) -> OrderBy -> SelectColumns -> SelectTable
    '    [Tip] 반환값을 변수에 저장하지 않고 실행만 할 때는 Set 키워드가 필요 없습니다.
    UnivDB.Query(T_SOURCE) _
          .Where("Dept", "Sales") _
          .Filter("Salary", ">=", 5000) _
          .OrderBy("Salary", True) _
          .SelectColumns("Name, Role, Salary") _
          .SelectTable() _
          .DumpToSheet Sheets(T_TARGET_SHEET).Range(T_TARGET_CELL), True
          
    MsgBox "데이터 추출 완료!"
End Sub

' ------------------------------------------------------------------------------
' [Template 2] 데이터 조작 (CRUD Operations)
' 설명: Insert, Update, Delete 메서드를 사용하여 데이터를 조작합니다.
' ------------------------------------------------------------------------------
Sub Template_02_CRUD()
    UnivDB.InitDB
    UnivDB.RegisterTable "Employees", Sheet1.Range("A1").CurrentRegion
    
    ' 1. [Insert] 신규 직원 추가
    '    SetValue로 컬럼값 설정 후 Insert 호출
    '    (Insert는 Long 타입을 반환하므로 Set을 사용하지 않습니다)
    UnivDB.Query("Employees") _
          .SetValue("ID", 100) _
          .SetValue("Name", "New Hire") _
          .SetValue("Dept", "HR") _
          .Insert
    
    ' 2. [Update] 부서 이동 및 승진 처리
    '    Where 조건에 맞는 행들의 값을 일괄 변경
    '    (Update는 Long 타입을 반환하므로 Set을 사용하지 않습니다)
    UnivDB.Query("Employees") _
          .Where("Name", "Bob") _
          .SetValue("Dept", "Finance") _
          .SetValue("Role", "Manager") _
          .Update
    
    ' 3. [Delete] 퇴사자 삭제
    '    조건에 맞는 행을 물리적으로 제거
    '    (Delete는 Long 타입을 반환하므로 Set을 사용하지 않습니다)
    UnivDB.Query("Employees") _
          .Where("Status", "Resigned") _
          .Delete
    
    ' 결과 확인을 위해 시트에 덤프
    UnivDB.GetTable("Employees").DumpToSheet Sheet1.Range("A1"), True
End Sub

' ------------------------------------------------------------------------------
' [Template 3] 리포트 오케스트레이션 (Report Builder Pipeline)
' 설명: clsReportBuilder를 사용하여 복잡한 데이터 가공 로직을 파이프라인 형태로 구성합니다.
'       (로드 -> 필터 -> 제외 -> 값수정 -> 컬럼선택 -> 출력)
' ------------------------------------------------------------------------------
Sub Template_03_ReportPipeline()
    UnivDB.InitDB
    UnivDB.RegisterTable "SalesData", Sheet1.Range("A1").CurrentRegion
    
    ' [Report Builder Pattern]
    ' UnivDB.Report 팩토리 메서드를 사용하여 변수 선언 없이 바로 체이닝 시작
    ' 순서: Init -> Exclude(삭제) -> Filter(선택) -> UpdateValue(수정) -> SelectFinal(컬럼) -> BuildToSheet(출력)
    UnivDB.Report("SalesData") _
      .Exclude("Status", "=", "Cancelled") _
      .Filter("Year", "=", 2024) _
      .UpdateValue("Grade", "VIP", "Amount", ">=", 10000) _
      .SelectFinal("Date", "Customer", "Amount", "Grade") _
      .DefineColumn("Date", vWidth:=12, vFormat:="yyyy-mm-dd", vAlign:="Center") _
      .DefineColumn("Amount", vWidth:=15, vFormat:="#,##0", vAlign:="Right") _
      .DefineColumn("Grade", vAlign:="Center", vHeaderAlias:="VIP Grade") _
      .BuildToSheet Sheet2, "A1"
    
    MsgBox "리포트 생성 완료!"
End Sub

' ------------------------------------------------------------------------------
' [Template 4] 고급 분석 (Join & Aggregation)
' 설명: 테이블 조인 후 집계(GroupBy)를 수행하는 분석 패턴입니다.
' ------------------------------------------------------------------------------
Sub Template_04_JoinAndAnalyze()
    Dim tJoined As clsTable
    Dim tSummary As clsTable
    
    UnivDB.InitDB
    UnivDB.RegisterTable "Orders", Sheet1.Range("A1").CurrentRegion
    UnivDB.RegisterTable "Customers", Sheet2.Range("A1").CurrentRegion
    
    ' 1. [Join] 주문(Left) + 고객(Right) 결합
    '    UnivDB.Join 파사드 사용
    Set tJoined = UnivDB.Join("Orders", "Customers", "CustID", jtLeft)
    
    ' 2. [GroupBy] 지역별 주문 금액 합계 및 평균 산출
    '    (clsTable의 GroupBy 기능 활용)
    Set tSummary = tJoined.GroupBy("Region", "Amount:SUM, Amount:AVG, OrderID:COUNT")
    
    ' 3. [Result] 결과 정렬 및 출력
    tSummary.OrderBy "Amount_SUM", True ' 내림차순 정렬
    tSummary.DumpToSheet Sheet3.Range("A1"), True
    
    MsgBox "분석 완료!"
End Sub

' ------------------------------------------------------------------------------
' [Template 5] 객체 직접 사용 (Direct Object Usage)
' 설명: RegisterTable 없이 clsTable 객체를 직접 생성하여 사용하는 패턴입니다.
'       일회성 작업이나 로컬 변수 스코프 관리가 필요할 때 유용합니다.
' ------------------------------------------------------------------------------
Sub Template_05_DirectObjectUsage()
    Dim tSource As clsTable
    Set tSource = UnivDB.CreateTable()
    
    ' 1. 객체 직접 로드 (RegisterTable 생략)
    '    전역 저장소(UnivDB)를 거치지 않으므로 가볍게 동작합니다.
    tSource.LoadFromSheet Sheet1.Range("A1").CurrentRegion
    
    ' 2. UnivDB 파사드에 객체 전달
    '    (문자열 이름 대신 tSource 객체 자체를 전달해도 완벽하게 호환됩니다)
    UnivDB.Query(tSource) _
          .Where("Dept", "IT") _
          .SelectColumns("Name, Dept") _
          .SelectTable() _
          .DumpToSheet Sheet2.Range("A1"), True
          
    MsgBox "객체 직접 전달 방식 완료!"
End Sub

' ------------------------------------------------------------------------------
' [Template 6] 연쇄 조인 및 마스터 동기화 (Cascade Join & Sync)
' 설명: 하위 단계(예: 계약)에서 상위 단계(공고->접수->요청)로 역방향 조인을 수행하여
'       완성된 데이터셋(Payload)을 만들고, 이를 Master DB에 일괄 반영(Upsert)합니다.
'       데이터 양이 적을 때(수천 건) 가장 효율적이고 안전한 패턴입니다.
' ------------------------------------------------------------------------------
Sub Template_06_CascadeJoinSync()
    ' [시나리오]
    ' 4단계(계약) -> 3단계(공고) -> 2단계(접수) -> 1단계(요청) 순으로 정보를 합쳐
    ' Master DB에 "계약번호"를 기준으로 덮어씁니다.
    
    UnivDB.InitDB
    ' (예시를 위해 현재 시트들을 각 단계 테이블로 등록한다고 가정)
    ' UnivDB.RegisterTable "Contract", Sheet4.Range("A1").CurrentRegion
    ' UnivDB.RegisterTable "Notice",   Sheet3.Range("A1").CurrentRegion
    ' UnivDB.RegisterTable "Receipt",  Sheet2.Range("A1").CurrentRegion
    ' UnivDB.RegisterTable "Request",  Sheet1.Range("A1").CurrentRegion
    ' UnivDB.RegisterTable "Master",   SheetMaster.Range("A1").CurrentRegion
    
    Dim T_Payload As clsTable
    
    ' 1. [Chain Join] 역방향 연쇄 조인 (Bottom-Up)
    '    Contract(4) + Notice(3) -> T_Step1
    '    T_Step1 + Receipt(2)    -> T_Step2
    '    T_Step2 + Request(1)    -> T_Payload (최종 완성)
    
    ' 4단계 + 3단계 (Key: 공고번호)
    Set T_Payload = UnivDB.Join("Contract", "Notice", "NoticeNo", jtLeft, "Not_")
    
    ' + 2단계 (Key: 접수번호)
    Set T_Payload = UnivDB.Join(T_Payload, "Receipt", "ReceiptNo", jtLeft, "Rec_")
    
    ' + 1단계 (Key: 조달요구번호)
    Set T_Payload = UnivDB.Join(T_Payload, "Request", "RequestNo", jtLeft, "Req_")
    
    ' 2. [Bulk Upsert] Master DB에 반영
    '    Master 테이블에 T_Payload의 데이터를 "ContractNo"(PK) 기준으로 병합합니다.
    '    (이미 존재하면 Update, 없으면 Insert)
    UnivDB.GetTable("Master").UpsertFrom T_Payload, "ContractNo"
    
    ' 3. 결과 확인
    ' UnivDB.GetTable("Master").DumpToSheet SheetMaster.Range("A1"), True
    
    MsgBox "마스터 DB 동기화 완료! (처리 건수: " & T_Payload.RowCount & "건)"
End Sub
