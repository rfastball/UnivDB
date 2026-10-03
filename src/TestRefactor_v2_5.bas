Attribute VB_Name = "TestRefactor_v2_5"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' ==============================================================================
' [Test Module] for UnivDB v2.5 Refactoring
' Description: Validates the JIT-optimized Filter engine and supporting functions.
' To Run: Call the main "RunAll_v2_5_Tests" sub.
' ==============================================================================

Private Const TEST_SHEET As String = "TestSheet"
Private ws As Worksheet

' ==============================================================================
' [Main Test Runner]
' ==============================================================================
Public Sub RunAll_v2_5_Tests()
    On Error GoTo TestFail
    
    Debug.Print "=================================================="
    Debug.Print "[[ Starting UnivDB v2.5 Refactor Tests... ]]"
    Debug.Print "=================================================="
    
    Test_Setup
    
    Debug.Print vbCrLf & "--- Testing Core Utils ---"
    Test_FastNorm
    Test_OrderBy_StrComp
    
    Debug.Print vbCrLf & "--- Testing JIT Filter Engine ---"
    Test_Filter_NumericPath
    Test_Filter_StringPath_Is
    Test_Filter_StringPath_Like
    
    Debug.Print vbCrLf & "=================================================="
    Debug.Print "[[ All v2.5 tests passed successfully! ]]"
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

Private Sub Test_FastNorm()
    Debug.Print "  - Running Test_FastNorm...";
    
    ' Test cases
    AssertEqual FastNorm("  abc "), "abc", "FastNorm_Trim"
    AssertEqual FastNorm(123.45), "123.45", "FastNorm_Number"
    AssertEqual FastNorm(Null), "", "FastNorm_Null"
    AssertEqual FastNorm(Empty), "", "FastNorm_Empty"
    AssertEqual FastNorm(CVErr(xlErrNA)), "", "FastNorm_Error"
    
    Debug.Print " Passed!"
End Sub

Private Sub Test_OrderBy_StrComp()
    Debug.Print "  - Running Test_OrderBy_StrComp...";
    
    ' Setup: Create data where numeric and string sorting differ
    Dim vData(1 To 4, 1 To 2)
    vData(1, 1) = "ID": vData(1, 2) = "Val"
    vData(2, 1) = 1: vData(2, 2) = "300" ' String
    vData(3, 1) = 2: vData(3, 2) = 50    ' Number
    vData(4, 1) = 3: vData(4, 2) = "1000" ' String
    
    Dim T As New clsTable
    T.LoadFromArray vData
    
    ' Test: Sort by the mixed-type 'Val' column
    T.OrderBy "Val", bDescending:=False ' Ascending sort
    
    Dim vResult As Variant
    vResult = T.RawArray
    
    ' Numeric types precede text; strings retain text ordering.
    AssertEqual CStr(vResult(2, 2)), "50", "OrderBy_Mixed_1"
    AssertEqual CStr(vResult(3, 2)), "1000", "OrderBy_Mixed_2"
    AssertEqual CStr(vResult(4, 2)), "300", "OrderBy_Mixed_3"
    Debug.Print " Passed!"
End Sub

Private Sub Test_Filter_NumericPath()
    Debug.Print "  - Running Test_Filter_NumericPath...";
    
    UnivDB.InitDB
    UnivDB.RegisterTable "Employees", ws.Range("A1").CurrentRegion
    
    Dim T As clsTable
    
    ' Test ">="
    Set T = UnivDB.Query("Employees").Filter("SALARY", ">=", 60000).SelectTable()
    AssertEqual T.RowCount - 1, 2, "Numeric_GTE" ' Bob, Charlie
    
    ' Test "<"
    Set T = UnivDB.Query("Employees").Filter("SALARY", "<", 50000).SelectTable()
    AssertEqual T.RowCount - 1, 1, "Numeric_LT" ' Alice
    
    ' Test "="
    Set T = UnivDB.Query("Employees").Filter("SALARY", "=", 75000).SelectTable()
    AssertEqual T.RowCount - 1, 1, "Numeric_EQ" ' Charlie
    
    Debug.Print " Passed!"
End Sub

Private Sub Test_Filter_StringPath_Is()
    Debug.Print "  - Running Test_Filter_StringPath_Is...";
    
    UnivDB.InitDB
    UnivDB.RegisterTable "Employees", ws.Range("A1").CurrentRegion
    
    Dim T As clsTable
    
    ' Test "IS" (case-insensitive equals)
    Set T = UnivDB.Query("Employees").Filter("DEPT", "IS", "sales").SelectTable()
    AssertEqual T.RowCount - 1, 1, "String_IS_Case"
    AssertEqual T.Value(2, "NAME"), "Bob", "String_IS_Value"
    
    ' Test "IS NOT"
    Set T = UnivDB.Query("Employees").Filter("DEPT", "IS NOT", "IT").SelectTable()
    AssertEqual T.RowCount - 1, 2, "String_IS_NOT"
    
    Debug.Print " Passed!"
End Sub

Private Sub Test_Filter_StringPath_Like()
    Debug.Print "  - Running Test_Filter_StringPath_Like...";
    
    UnivDB.InitDB
    UnivDB.RegisterTable "Employees", ws.Range("A1").CurrentRegion
    
    Dim T As clsTable
    
    ' Test "LIKE" (case-insensitive)
    Set T = UnivDB.Query("Employees").Filter("NAME", "LIKE", "c*").SelectTable()
    AssertEqual T.RowCount - 1, 1, "String_LIKE_Case"
    AssertEqual T.Value(2, "NAME"), "Charlie", "String_LIKE_Value"
    
    ' Test "NOT LIKE"
    Set T = UnivDB.Query("Employees").Filter("STATUS", "NOT LIKE", "Active*").SelectTable()
    AssertEqual T.RowCount - 1, 1, "String_NOT_LIKE"
    AssertEqual T.Value(2, "NAME"), "David", "String_NOT_LIKE_Value"
    
    Debug.Print " Passed!"
End Sub


' ==============================================================================
' [2] Setup and Teardown
' ==============================================================================

Private Sub Test_Setup()
    ' Add a new worksheet for testing
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(TEST_SHEET).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = Worksheets.Add
    ws.Name = TEST_SHEET
    
    ' Create sample data
    Dim vData(1 To 5, 1 To 4)
    vData(1, 1) = "ID": vData(1, 2) = "NAME": vData(1, 3) = "DEPT": vData(1, 4) = "SALARY"
    vData(2, 1) = 101: vData(2, 2) = "Alice": vData(2, 3) = "IT": vData(2, 4) = 48000
    vData(3, 1) = 102: vData(3, 2) = "Bob": vData(3, 3) = "Sales": vData(3, 4) = 60000
    vData(4, 1) = 103: vData(4, 2) = "Charlie": vData(4, 3) = "IT": vData(4, 4) = 75000
    vData(5, 1) = 104: vData(5, 2) = "David": vData(5, 3) = "HR": vData(5, 4) = 52000
    
    ' Add a status column for NOT LIKE test
    ws.Range("E1").Value = "STATUS"
    ws.Range("E2").Value = "Active"
    ws.Range("E3").Value = "Active"
    ws.Range("E4").Value = "Active"
    ws.Range("E5").Value = "Inactive"
    
    ws.Range("A1").Resize(5, 4).Value = vData
    ws.Columns.AutoFit
End Sub

Private Sub Test_Teardown()
    ' Clean up the test sheet
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        Set ws = Nothing
    End If
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
