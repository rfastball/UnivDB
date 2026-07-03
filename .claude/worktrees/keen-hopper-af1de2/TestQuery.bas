Attribute VB_Name = "TestQuery"
Option Explicit

' ==============================================================================
' Test Module for clsQuery (Insert, Update, Delete)
' 사용법: VBE에서 RunAllTests 서브루틴을 실행하세요.
' ==============================================================================

Public Sub RunAllTests()
    Debug.Print "=== Starting Tests ==="
    Test_Insert
    Test_Update
    Test_Delete
    Test_ComplexScenario
    Debug.Print "=== All Tests Completed ==="
End Sub

' 테스트용 더미 테이블 생성 헬퍼
Private Function GetDummyTable() As clsTable
    Dim T As clsTable
    Set T = UnivDB.CreateTable()
    Dim arrData(1 To 4, 1 To 3) As Variant
    
    ' Header
    arrData(1, 1) = "ID"
    arrData(1, 2) = "Name"
    arrData(1, 3) = "Dept"
    
    ' Data
    arrData(2, 1) = 1: arrData(2, 2) = "Alice": arrData(2, 3) = "HR"
    arrData(3, 1) = 2: arrData(3, 2) = "Bob":   arrData(3, 3) = "IT"
    arrData(4, 1) = 3: arrData(4, 2) = "Carol": arrData(4, 3) = "Sales"
    
    T.LoadFromArray arrData
    Set GetDummyTable = T
End Function

Sub Test_Insert()
    Debug.Print "[Test_Insert] Running..."
    Dim T As clsTable
    Dim lCount As Long
    
    Set T = GetDummyTable()
    
    ' Insert New Row
    lCount = UnivDB.Query(T) _
              .SetValue("ID", 4) _
              .SetValue("Name", "Dave") _
              .SetValue("Dept", "IT") _
              .Insert()
              
    ' Assertions
    Debug.Assert lCount = 1
    Debug.Assert T.RowCount = 5             ' Header(1) + Data(3) + New(1) = 5
    Debug.Assert T.Value(5, "Name") = "Dave" ' 5번째 행에 추가됨
    Debug.Assert T.Value(5, "Dept") = "IT"
    
    Debug.Print "[Test_Insert] Passed."
End Sub

Sub Test_Update()
    Debug.Print "[Test_Update] Running..."
    Dim T As clsTable
    Dim lCount As Long
    
    Set T = GetDummyTable()
    
    ' Update Bob's Dept to 'Finance'
    lCount = UnivDB.Query(T) _
              .Where("Name", "Bob") _
              .SetValue("Dept", "Finance") _
              .Update()
              
    ' Assertions
    Debug.Assert lCount = 1
    Debug.Assert T.Value(3, "Dept") = "Finance" ' Bob is at Row 3 (Header=1, Alice=2, Bob=3)
    Debug.Assert T.Value(2, "Dept") = "HR"      ' Alice is at Row 2
    
    Debug.Print "[Test_Update] Passed."
End Sub

Sub Test_Delete()
    Debug.Print "[Test_Delete] Running..."
    Dim T As clsTable
    Dim lCount As Long
    
    Set T = GetDummyTable()
    
    ' Delete Carol (Sales)
    lCount = UnivDB.Query(T) _
              .Where("Dept", "Sales") _
              .Delete()
              
    ' Assertions
    Debug.Assert lCount = 1
    Debug.Assert T.RowCount = 3 ' Header(1) + Alice(1) + Bob(1) = 3 (Carol Deleted)
    
    ' Verify remaining data (Alice and Bob)
    Debug.Assert T.Value(2, "Name") = "Alice" ' Row 2
    Debug.Assert T.Value(3, "Name") = "Bob"   ' Row 3
    
    Debug.Print "[Test_Delete] Passed."
End Sub

Sub Test_ComplexScenario()
    Debug.Print "[Test_ComplexScenario] Running..."
    Dim T As clsTable
    
    Set T = GetDummyTable()
    
    ' 1. Insert Dave (ID=4, IT)
    UnivDB.Query(T).SetValue("ID", 4).SetValue("Name", "Dave").SetValue("Dept", "IT").Insert
    
    ' 2. Update all IT dept to 'Dev' (Bob & Dave)
    UnivDB.Query(T).Where("Dept", "IT").SetValue("Dept", "Dev").Update
    
    ' 3. Delete HR (Alice)
    UnivDB.Query(T).Where("Dept", "HR").Delete
    
    ' Expected Result:
    ' Alice (HR) -> Deleted
    ' Bob (IT) -> Dev (Shifted to Row 1)
    ' Carol (Sales) -> Sales (Shifted to Row 2)
    ' Dave (IT) -> Dev (Shifted to Row 3)
    
    Debug.Assert T.RowCount = 4                 ' Header(1) + Bob + Carol + Dave = 4
    Debug.Assert T.Value(2, "Dept") = "Dev"     ' Bob (Row 2)
    Debug.Assert T.Value(3, "Dept") = "Sales"   ' Carol (Row 3)
    Debug.Assert T.Value(4, "Dept") = "Dev"     ' Dave (Row 4)
    
    Debug.Print "[Test_ComplexScenario] Passed."
End Sub
