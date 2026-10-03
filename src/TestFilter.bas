Attribute VB_Name = "TestFilter"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' ==============================================================================
' [Test Module] UnivDB Filter Unit Tests
' 설명: clsQuery의 연산자별(수치 vs 문자열) 필터링 동작을 정밀 검증합니다.
' ==============================================================================

Public Sub Run_Filter_Unit_Tests()
    On Error GoTo TestFail
    
    Debug.Print "=== [UnivDB Filter Test 시작] ==="
    
    ' 1. 테스트 데이터 준비
    Setup_Test_Data
    
    ' 2. 수치 연산자 테스트 (Numeric Path)
    Test_Numeric_Filters
    
    ' 3. 문자열 연산자 테스트 (String Path)
    Test_String_Filters
    
    ' 4. 복합 체이닝 테스트 (Chained Filters)
    Test_Chained_Filters

    Debug.Print "=== [모든 필터 테스트 통과!] ==="
    Exit Sub

TestFail:
    Debug.Print "!!! 테스트 실패: " & Err.Description
End Sub

' ------------------------------------------------------------------------------
' [Setup] 테스트용 가상 테이블 등록
' ------------------------------------------------------------------------------
Private Sub Setup_Test_Data()
    Dim vData(1 To 6, 1 To 3) As Variant
    
    ' 헤더
    vData(1, 1) = "ID": vData(1, 2) = "Amount": vData(1, 3) = "Dept"
    
    ' 데이터 (수치와 문자열 혼합)
    vData(2, 1) = 1: vData(2, 2) = 1000:   vData(2, 3) = "IT"
    vData(3, 1) = 2: vData(3, 2) = 0:      vData(3, 3) = "Sales"
    vData(4, 1) = 3: vData(4, 2) = 5000:   vData(4, 3) = "MP"
    vData(5, 1) = 4: vData(5, 2) = 2000:   vData(5, 3) = "MD"
    vData(6, 1) = 5: vData(6, 2) = "N/A":  vData(6, 3) = "MW" ' 수치 열에 문자 포함 케이스
    
    UnivDB.InitDB
    UnivDB.RegisterTable "T_Test", vData
End Sub

' ------------------------------------------------------------------------------
' [Test 1] 수치 연산자 검증 (=, <>, >, <, >=, <=)
' ------------------------------------------------------------------------------
Private Sub Test_Numeric_Filters()
    Dim T As Object
    Debug.Print "  - 수치 필터 테스트 중..."
    
    ' 1-1. 금액이 0이 아닌 것 (<>) : 1000, 5000, 2000 추출 예상 (3건)
    ' [주의] "N/A"는 IsNumeric이 False이므로 제외됨
    Set T = UnivDB.Query("T_Test").Filter("Amount", "<>", 0).SelectTable()
    AssertEqual T.RowCount - 1, 3, "Numeric_NotEqual_Zero"
    
    ' 1-2. 금액이 2000 이상 (>=) : 5000, 2000 추출 예상 (2건)
    Set T = UnivDB.Query("T_Test").Filter("Amount", ">=", 2000).SelectTable()
    AssertEqual T.RowCount - 1, 2, "Numeric_GTE"
    
    Debug.Print "    OK."
End Sub

' ------------------------------------------------------------------------------
' [Test 2] 문자열 연산자 검증 (IS, IS NOT, LIKE, NOT LIKE)
' ------------------------------------------------------------------------------
Private Sub Test_String_Filters()
    Dim T As Object
    Debug.Print "  - 문자열 필터 테스트 중..."
    
    ' 2-1. 대소문자 무시 일치 (IS)
    Set T = UnivDB.Query("T_Test").Filter("Dept", "IS", "it").SelectTable()
    AssertEqual T.RowCount - 1, 1, "String_IS"
    
    ' 2-2. 특정 문자 제외 (IS NOT)
    Set T = UnivDB.Query("T_Test").Filter("Dept", "IS NOT", "MP").SelectTable()
    AssertEqual T.RowCount - 1, 4, "String_IS_NOT"
    
    ' 2-3. 패턴 매칭 (LIKE) - [수정됨] 주석 기호(') 추가
    Set T = UnivDB.Query("T_Test").Filter("Dept", "LIKE", "M*").SelectTable()
    AssertEqual T.RowCount - 1, 3, "String_LIKE" ' (MP, MD, MW) 가 정상 추출되어야 함
    
    ' 2-4. 패턴 제외 (NOT LIKE) - [수정됨] 주석 기호(') 추가
    Set T = UnivDB.Query("T_Test").Filter("Dept", "NOT LIKE", "*S*").SelectTable()
    AssertEqual T.RowCount - 1, 4, "String_NOT_LIKE" ' (Sales 제외)
    
    Debug.Print "    OK."
End Sub

' ------------------------------------------------------------------------------
' [Test 3] 필터 체이닝 (AND 조건)
' ------------------------------------------------------------------------------
Private Sub Test_Chained_Filters()
    Dim T As Object
    Debug.Print "  - 체이닝 필터 테스트 중..."
    
    ' 금액 <> 0 AND 부서 NOT LIKE M*
    ' 대상: 1(IT/1000), 2(Sales/0), 3(MP/5000), 4(MD/2000), 5(MW/NA)
    ' 결과: IT(1000) 1건만 남아야 함
    Set T = UnivDB.Query("T_Test") _
            .Filter("Amount", "<>", 0) _
            .Filter("Dept", "NOT LIKE", "M*") _
            .SelectTable()
            
    AssertEqual T.RowCount - 1, 1, "Chained_Filter_Logic"
    AssertEqual T.Value(2, "Dept"), "IT", "Chained_Filter_Value"
    
    Debug.Print "    OK."
End Sub

' ------------------------------------------------------------------------------
' [Helper] 단언문 (Assertion)
' ------------------------------------------------------------------------------
Private Sub AssertEqual(ByVal actual As Variant, ByVal expected As Variant, ByVal testName As String)
    If CStr(actual) <> CStr(expected) Then
        Err.Raise 9999, "AssertEqual", "실패 [" & testName & "] -> 기대값: " & expected & " / 실제값: " & actual
    End If
End Sub
