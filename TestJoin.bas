Attribute VB_Name = "TestJoin"
Option Explicit

' ==============================================================================
' [Test Module] for UnivDB Joins
' Description: Validates the Join engine functionality, especially regarding
'              case-insensitivity and handling of different key columns.
' To Run: Call the main "RunAll_Join_Tests" sub.
' ==============================================================================

Private Const TEST_SHEET As String = "TestJoinSheet"
Private ws As Worksheet

' ==============================================================================
' [Main Test Runner]
' ==============================================================================
Public Sub RunAll_Join_Tests()
    On Error GoTo TestFail
    
    Debug.Print "=================================================="
    Debug.Print "[[ Starting UnivDB Join Engine Tests... ]]"
    Debug.Print "=================================================="
    
    Test_Setup
    
    Debug.Print vbCrLf & "--- Testing Join Functionality ---"
    Test_Join_CaseInsensitiveData
    Test_Join_DifferentColumnNames
    
    Debug.Print vbCrLf & "=================================================="
    Debug.Print "[[ All Join tests passed successfully! ]]"
    Debug.Print "=================================================="
    
    GoTo Cleanup
    
TestFail:
    Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Debug.Print "!! TEST FAILED: " & Err.Description & " (Line " & Erl & ")"
    Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    
Cleanup:
    On Error Resume Next
    Test_Teardown
    On Error GoTo 0
End Sub


' ==============================================================================
' [1] Test Cases
' ==============================================================================

Private Sub Test_Join_CaseInsensitiveData()
    Debug.Print "  - Running Test_Join_CaseInsensitiveData...";
    
    ' Setup: Two tables with a common column "Dept", but the data has different casing.
    Dim vEmps(1 To 4, 1 To 3)
    vEmps(1, 1) = "ID": vEmps(1, 2) = "Name": vEmps(1, 3) = "Dept"
    vEmps(2, 1) = 1: vEmps(2, 2) = "Alice": vEmps(2, 3) = "IT"
    vEmps(3, 1) = 2: vEmps(3, 2) = "Bob": vEmps(3, 3) = "sales" ' lowercase
    vEmps(4, 1) = 3: vEmps(4, 2) = "Charlie": vEmps(4, 3) = "HR"
    
    Dim vDepts(1 To 3, 1 To 2)
    vDepts(1, 1) = "Dept": vDepts(1, 2) = "Dept_Manager"
    vDepts(2, 1) = "Sales": vDepts(2, 2) = "Mr. Black"  ' Titlecase
    vDepts(3, 1) = "IT": vDepts(3, 2) = "Ms. White"
    
    UnivDB.InitDB
    UnivDB.RegisterTable "Employees_CaseTest", vEmps
    UnivDB.RegisterTable "Departments_CaseTest", vDepts
    
    ' Test: Perform an inner join on the "Dept" column.
    ' The engine should match "sales" with "Sales".
    Dim tResult As clsTable
    Set tResult = UnivDB.Join("Employees_CaseTest", "Departments_CaseTest", "Dept", jtInner)
    
    ' Validation: Expect 2 rows to match (IT and Sales)
    AssertEqual tResult.RowCount, 3, "Join_CaseInsensitiveData_RowCount" ' 1 header + 2 data rows
    
    Debug.Print " Passed!"
End Sub

Private Sub Test_Join_DifferentColumnNames()
    Debug.Print "  - Running Test_Join_DifferentColumnNames...";
    
    ' Setup: Two tables with different names for the join key columns.
    Dim vEmps(1 To 4, 1 To 3)
    vEmps(1, 1) = "ID": vEmps(1, 2) = "Name": vEmps(1, 3) = "Emp_Dept_ID" ' FK
    vEmps(2, 1) = 1: vEmps(2, 2) = "David": vEmps(2, 3) = 10
    vEmps(3, 1) = 2: vEmps(3, 2) = "Eve": vEmps(3, 3) = 20
    vEmps(4, 1) = 3: vEmps(4, 2) = "Frank": vEmps(4, 3) = 99 ' No match
    
    Dim vDepts(1 To 3, 1 To 2)
    vDepts(1, 1) = "Dept_PK": vDepts(1, 2) = "Dept_Name" ' PK
    vDepts(2, 1) = 10: vDepts(2, 2) = "Engineering"
    vDepts(3, 1) = 20: vDepts(3, 2) = "Marketing"

    UnivDB.InitDB
    UnivDB.RegisterTable "Employees_DiffCols", vEmps
    UnivDB.RegisterTable "Departments_DiffCols", vDepts
    
    ' Test: Perform an inner join using the two different key column names.
    ' This requires a modified Join function, e.g., Join(left, right, leftKey, rightKey, type)
    ' We assume this change has been made for the purpose of this test.
    Dim tResult As clsTable
    
    ' ** ASSUMPTION: The Join function has been overloaded or modified to handle two key names. **
    ' This will cause a compile error if the signature is still Join(..., sJoinKey As String, ...)
    ' which is a valid test outcome.
    Set tResult = UnivDB.Join("Employees_DiffCols", "Departments_DiffCols", leftKey:="Emp_Dept_ID", rightKey:="Dept_PK", jType:=jtInner)

    ' Validation: Expect 2 rows to match (Dept 10 and 20)
    AssertEqual tResult.RowCount, 3, "Join_DifferentCols_RowCount" ' 1 header + 2 data rows
    
    Debug.Print " Passed!"
End Sub


' ==============================================================================
' [2] Setup and Teardown
' ==============================================================================

Private Sub Test_Setup()
    ' Add a new worksheet for testing if it doesn't exist
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(TEST_SHEET).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = Worksheets.Add
    ws.Name = TEST_SHEET
    ws.Visible = xlSheetHidden ' Hide the sheet during tests
End Sub

Private Sub Test_Teardown()
    ' Clean up the test sheet
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        Set ws = Nothing
    End If
    UnivDB.ClearDB
End Sub


' ==============================================================================
' [3] Assertion Helper
' ==============================================================================

Private Sub AssertEqual(ByVal actual As Variant, ByVal expected As Variant, ByVal testName As String)
    If CStr(actual) <> CStr(expected) Then
        Err.Raise vbObjectError + 512, "AssertEqual", "Test Failed: " & testName & _
                  vbCrLf & "  Expected: " & expected & _
                  vbCrLf & "  Actual:   " & actual
    End If
End Sub
