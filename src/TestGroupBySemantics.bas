Attribute VB_Name = "TestGroupBySemantics"
' @desc GroupBy 키 의미론(대소문자, Empty, 복합 키 충돌) 회귀 검증 담당
' @depends UnivDB
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' Phase R1 (결정 B 적용됨): GroupBy key 의미론은 인덱스 계약(modIndexKey)과 동일하다.
' vbTextCompare(대소문자 무시) / Empty->"" / Error 키 행 skip.
' 남은 한계: 복합 키 CStr & "|" 연결의 구분자 충돌 (R3에서 encode helper로 해소 예정).

' 결정 B: "IT"와 "it"는 GroupBy에서도 인덱스와 같은 한 키. 그룹 라벨은 첫 등장 표기
Public Sub Test_GroupBySem_CaseKey()
    Dim T As clsTable
    Dim tG As clsTable

    Set T = NewCaseTable()

    Set tG = T.GroupBy("Dept", "Score:SUM")
    LogShim.AssertEq "gbsem_case_groups", 2, tG.RowCount ' 헤더 + 통합 1그룹
    LogShim.AssertEq "gbsem_case_first_label", "IT", tG.Value(2, "Dept")
    LogShim.AssertEq "gbsem_case_first_sum", 35, tG.Value(2, "Score_SUM")

    ' 패리티: 인덱스 경로와 같은 키 의미론
    LogShim.AssertEq "gbsem_case_index_rows", 3, T.SelectRows("Dept", "it").Count
End Sub

' 결정 B: 그룹 키가 Error인 행은 그룹핑에서 제외 (인덱스 AddRow의 Error skip과 동치)
Public Sub Test_GroupBySem_ErrorKeySkip()
    Dim T As clsTable
    Dim tG As clsTable
    Dim vData(1 To 4, 1 To 2) As Variant

    vData(1, 1) = "Dept": vData(1, 2) = "Score"
    vData(2, 1) = "IT": vData(2, 2) = 10
    vData(3, 1) = CVErr(xlErrNA): vData(3, 2) = 20
    vData(4, 1) = "IT": vData(4, 2) = 30

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    Set tG = T.GroupBy("Dept", "Score:SUM")
    LogShim.AssertEq "gbsem_errkey_groups", 2, tG.RowCount ' 헤더 + IT (Error 행 제외)
    LogShim.AssertEq "gbsem_errkey_sum", 40, tG.Value(2, "Score_SUM")
End Sub

' 결정 B: Empty 그룹 값은 인덱스 계약대로 "" 키로 묶임 (NormalizeIndexKey)
Public Sub Test_GroupBySem_EmptyKey()
    Dim T As clsTable
    Dim tG As clsTable
    Dim vData(1 To 4, 1 To 2) As Variant

    vData(1, 1) = "Dept": vData(1, 2) = "Score"
    vData(2, 1) = "IT": vData(2, 2) = 10
    ' vData(3, 1) 미할당 = Empty
    vData(3, 2) = 20
    vData(4, 1) = "IT": vData(4, 2) = 30

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    Set tG = T.GroupBy("Dept", "Score:SUM")
    LogShim.AssertEq "gbsem_empty_groups", 3, tG.RowCount ' 헤더 + IT + ""
    LogShim.AssertEq "gbsem_empty_label", "", tG.Value(3, "Dept")
    LogShim.AssertEq "gbsem_empty_sum", 20, tG.Value(3, "Score_SUM")
End Sub

' R3: 복합 키는 길이 접두 인코딩(modIndexKey.EncodeCompositeKeyPart)이라
' 구분자 포함 값도 충돌하지 않는다 - ("X|Y","Z")와 ("X","Y|Z")는 다른 그룹
Public Sub Test_GroupBySem_CompositeDelimiterCollision()
    Dim T As clsTable
    Dim tG As clsTable
    Dim vData(1 To 3, 1 To 3) As Variant

    vData(1, 1) = "A": vData(1, 2) = "B": vData(1, 3) = "Score"
    vData(2, 1) = "X|Y": vData(2, 2) = "Z": vData(2, 3) = 10
    vData(3, 1) = "X": vData(3, 2) = "Y|Z": vData(3, 3) = 20

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    Set tG = T.GroupBy("A, B", "Score:SUM")
    LogShim.AssertEq "gbsem_collision_groups", 3, tG.RowCount ' 헤더 + 구분된 2그룹
    LogShim.AssertEq "gbsem_collision_first_a", "X|Y", tG.Value(2, "A")
    LogShim.AssertEq "gbsem_collision_first_sum", 10, tG.Value(2, "Score_SUM")
    LogShim.AssertEq "gbsem_collision_second_b", "Y|Z", tG.Value(3, "B")
    LogShim.AssertEq "gbsem_collision_second_sum", 20, tG.Value(3, "Score_SUM")
End Sub

' 복합 키 경로(누적기)도 단일 컬럼 엔진 경로와 같은 키 계약을 따른다:
' 대소문자 무시 병합(라벨은 첫 등장 표기) + Error 키 행 skip
Public Sub Test_GroupBySem_CompositeContractParity()
    Dim T As clsTable
    Dim tG As clsTable
    Dim vData(1 To 4, 1 To 3) As Variant

    vData(1, 1) = "Dept": vData(1, 2) = "Tag": vData(1, 3) = "Score"
    vData(2, 1) = "IT": vData(2, 2) = "x": vData(2, 3) = 10
    vData(3, 1) = "it": vData(3, 2) = "X": vData(3, 3) = 20
    vData(4, 1) = CVErr(xlErrNA): vData(4, 2) = "x": vData(4, 3) = 99

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    Set tG = T.GroupBy("Dept, Tag", "Score:SUM")
    LogShim.AssertEq "gbsem_comp_groups", 2, tG.RowCount ' 헤더 + 통합 1그룹 (Error 행 제외)
    LogShim.AssertEq "gbsem_comp_label_dept", "IT", tG.Value(2, "Dept")
    LogShim.AssertEq "gbsem_comp_label_tag", "x", tG.Value(2, "Tag")
    LogShim.AssertEq "gbsem_comp_sum", 30, tG.Value(2, "Score_SUM")
End Sub

' 현행 유지: 그룹 컬럼 값은 키 문자열 재분해(Split)로 복원되므로 숫자도 String으로 나온다
' (타입 보존은 비결정 D(NULL/타입 의미론) 하위 항목 - R1 스코프 아님)
Public Sub Test_GroupBySem_NumericKeyType()
    Dim T As clsTable
    Dim tG As clsTable
    Dim vData(1 To 4, 1 To 2) As Variant

    vData(1, 1) = "Code": vData(1, 2) = "Score"
    vData(2, 1) = 7: vData(2, 2) = 10
    vData(3, 1) = 7: vData(3, 2) = 20
    vData(4, 1) = 8: vData(4, 2) = 30

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    Set tG = T.GroupBy("Code", "Score:SUM")
    LogShim.AssertEq "gbsem_numkey_groups", 3, tG.RowCount
    LogShim.AssertEq "gbsem_numkey_type", "String", TypeName(tG.Value(2, "Code"))
    LogShim.AssertEq "gbsem_numkey_value", "7", tG.Value(2, "Code")
End Sub

Private Function NewCaseTable() As clsTable
    Dim vData(1 To 4, 1 To 2) As Variant
    Dim T As clsTable

    vData(1, 1) = "Dept": vData(1, 2) = "Score"
    vData(2, 1) = "IT": vData(2, 2) = 10
    vData(3, 1) = "it": vData(3, 2) = 5
    vData(4, 1) = "IT": vData(4, 2) = 20

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData
    Set NewCaseTable = T
End Function
