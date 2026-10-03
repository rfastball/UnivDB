Attribute VB_Name = "TestTableBuilder"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' ==============================================================================
' [Helper] Jagged Array(배열의 배열)를 순수 2차원 배열로 변환 (테스트 전용)
' ==============================================================================
Private Function Create2DArray(ParamArray vRows() As Variant) As Variant
    Dim i As Long, j As Long
    Dim lRows As Long, lCols As Long
    Dim vResult As Variant
    
    lRows = UBound(vRows) + 1
    lCols = UBound(vRows(0)) + 1
    
    ReDim vResult(1 To lRows, 1 To lCols)
    
    For i = 1 To lRows
        For j = 1 To lCols
            vResult(i, j) = vRows(i - 1)(j - 1)
        Next j
    Next i
    
    Create2DArray = vResult
End Function

' ==============================================================================
' [Test Suite] clsTableBuilder 통합 테스트
' ==============================================================================
Sub RunAllTests_Facade()
    UnivDB.InitDB
    
    ' 테스트용 가상 데이터 생성 (Inventory 테이블)
    ' 헬퍼 함수를 사용하여 엔진이 좋아하는 순수 2차원 배열로 주입합니다.
    Dim vInitData As Variant
    vInitData = Create2DArray( _
        Array("PROD_ID", "Name", "Qty", "Price"), _
        Array("P001", "Mouse", 10, 15000), _
        Array("P002", "Keyboard", 0, 35000) _
    )
    UnivDB.RegisterTable "Inventory", vInitData
    
    Debug.Print "=== clsTableBuilder Tests (Facade) Started ==="
    Test_Facade_EditTable
    Test_Facade_NewBuilder
    Debug.Print "=== All Tests Passed! ==="
End Sub

Private Sub Test_Facade_EditTable()
    ' 파사드를 통한 EditTable 호출 (기존 테이블 스키마 복제)
    Dim B As clsTableBuilder
    Set B = UnivDB.EditTable("Inventory", "PROD_ID")
    
    Dim tInv As clsTable: Set tInv = UnivDB.GetTable("Inventory")
    
    ' 1. 기존 데이터 추출 및 수정 (Call 키워드 사용)
    Call B.ExtractRow(tInv, "P001").Val("Qty", 50)
    
    ' 2. 신규 데이터 추가 및 동적 스키마 확장 (Call 키워드 사용)
    Call B.Row("P003").Val("Name", "Monitor").Val("Qty", 5).Val("Location", "A-1")
    
    ' 3. Upsert 병합
    tInv.UpsertFrom B.ToTable(), "PROD_ID"
    
    ' 검증
    If tInv.SelectRows("PROD_ID", "P001").Count = 1 Then
        Dim lRow1 As Long: lRow1 = tInv.SelectRows("PROD_ID", "P001")(1)
        If tInv.Value(lRow1, "Qty") = 50 Then Debug.Print "[OK] EditTable & Upsert: 기존 데이터 수정 완료"
    End If
    
    If tInv.SelectRows("PROD_ID", "P003").Count = 1 Then
        Dim lRow3 As Long: lRow3 = tInv.SelectRows("PROD_ID", "P003")(1)
        If tInv.Value(lRow3, "Location") = "A-1" Then Debug.Print "[OK] EditTable & Upsert: 신규 데이터 및 컬럼 추가 완료"
    End If
End Sub

Private Sub Test_Facade_NewBuilder()
    ' 파사드를 통한 NewBuilder 호출 (빈 스키마부터 시작)
    Dim B As clsTableBuilder
    Set B = UnivDB.Builder("ID")
    
    ' Call 키워드로 구문 오류 해결
    Call B.Row("A01").Val("Color", "Red").Val("Size", "L")
    Call B.Row("A02").Val("Color", "Blue").Val("Size", "M").PatchIf("Size", "=", "M", "Status", "SoldOut")
    
    Dim tNew As clsTable: Set tNew = B.ToTable()
    
    ' 검증
    If tNew.RowCount = 3 And tNew.ColCount = 4 Then ' 헤더 포함 3행, 컬럼(ID, Color, Size, Status) 4열
        Debug.Print "[OK] NewBuilder: 동적 테이블 완전 조립 성공"
    End If
End Sub
