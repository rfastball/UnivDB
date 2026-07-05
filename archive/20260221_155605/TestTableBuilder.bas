Attribute VB_Name = "TestTableBuilder"
Option Explicit

' ==============================================================================
' [Test] clsTableBuilder Standard Workflow
' ==============================================================================
Public Sub Test_BuilderWorkflow()
    ' 0. 테스트 환경 초기화
    UnivDB.ClearDB
    UnivDB.InitDB
    
    ' 1. 원본 "Inventory" 테이블 생성 및 등록
    Dim vInventory As Variant
    ReDim vInventory(1 To 4, 1 To 5)
    
    ' 헤더
    vInventory(1, 1) = "PROD_ID"
    vInventory(1, 2) = "ProductName"
    vInventory(1, 3) = "Qty"
    vInventory(1, 4) = "Price"
    vInventory(1, 5) = "Status"
    
    ' 데이터
    vInventory(2, 1) = "PROD_001": vInventory(2, 2) = "Old Monitor": vInventory(2, 3) = 10: vInventory(2, 4) = 150: vInventory(2, 5) = "In Stock"
    vInventory(3, 1) = "PROD_003": vInventory(3, 2) = "Mouse": vInventory(3, 3) = 200: vInventory(3, 4) = 25: vInventory(3, 5) = "In Stock"
    vInventory(4, 1) = "PROD_004": vInventory(4, 2) = "Webcam": vInventory(4, 3) = 0: vInventory(4, 4) = 80: vInventory(4, 5) = "In Stock"

    UnivDB.RegisterTable "Inventory", vInventory
    
    Debug.Print "--- [Before] Original Inventory Table ---"
    UnivDB.GetTable("Inventory").Query.DebugPrint
    
    ' ==============================================================================
    ' [WORKFLOW START]
    ' ==============================================================================
    
    Dim B As clsTableBuilder
    ' 1. 진입: "Inventory" 테이블의 구조를 본뜬 작업대 생성 (기준키: "PROD_ID")
    Set B = UnivDB.EditTable("Inventory", "PROD_ID")
    
    ' 2. 절차적 가공: 특정 행들을 콕 집어서(ID 기반) 값 수정
    
    ' [Case A] 기존 데이터 추출 후 조작 (PROD_001)
    ' Status를 "Out of Stock"으로 변경하는 조건에 걸리는지 확인하기 위해 Qty를 0으로 변경
    B.ExtractRow(UnivDB.GetTable("Inventory"), "PROD_001") _
     .Val "Qty", 0 _
     .Val "LastUpdated", Now _
     .PatchIf "Qty", "<=", 0, "Status", "Out of Stock"
     
    ' [Case B] 테이블에 없는 신규 데이터 조립 (PROD_002)
    B.Row("PROD_002") _
     .Val "ProductName", "New Keyboard" _
     .Val "Qty", 100 _
     .Val "Remarks", "Initial Stock" ' 기존 테이블에 없던 'Remarks' 컬럼 자동 추가
     
    ' [Case C] 현재 조립 중인 값을 읽어서(Getter) 계산에 활용 (PROD_003)
    B.Row "PROD_003"
    B.Val "TotalValue", B.Val("Qty") * B.Val("Price")
    
    Debug.Print vbCrLf & "--- [Staging] Builder's internal data ---"
    B.ToTable.Query.DebugPrint ' Upsert 전에 빌더 내용 확인
    
    ' 3. 반영: 작업대의 데이터를 clsTable로 변환 후, 원본 테이블에 일괄 병합 (Upsert)
    UnivDB.GetTable("Inventory").UpsertFrom B.ToTable(), "PROD_ID"
    
    ' ==============================================================================
    ' [WORKFLOW END]
    ' ==============================================================================
    
    Debug.Print vbCrLf & "--- [After] Merged Inventory Table ---"
    UnivDB.GetTable("Inventory").Query.DebugPrint
    
    ' 검증:
    ' 1. PROD_001: Qty=0, Status="Out of Stock", LastUpdated에 날짜 찍힘
    ' 2. PROD_002: 신규 행으로 추가됨, Remarks 컬럼 추가됨
    ' 3. PROD_003: TotalValue 컬럼 추가되고 값(200*25=5000) 계산됨
    ' 4. PROD_004: 원본 데이터 그대로 유지
    Stop ' VBE에서 중단하고 직접 데이터 확인
    
    MsgBox "개별 레코드 조립 및 병합 완료!"

End Sub

