Attribute VB_Name = "TestGroupBySemantics"
' @desc GroupBy 키 의미론(대소문자, Empty, 복합 키 충돌) 회귀 검증 담당
' @depends UnivDB
Option Explicit

' Phase R1: GroupBy key 의미론 확정 전, 현행 동작을 그대로 고정하는 회귀 스위트.
' 현행: 로컬 Dictionary 누적기 CompareMode 미설정 = binary(대소문자 구분).
' 인덱스 계약(modIndexKey: vbTextCompare, Empty->"", Error skip)과 다르다.
' 결정 B(계약 통일) 적용 시 이 파일의 기대값 변경이 별도 커밋으로 뒤따른다.

' 현행: "IT"와 "it"는 GroupBy에서 다른 그룹, 인덱스(Where/SelectRows)에서는 같은 키
Public Sub Test_GroupBySem_CaseKey()
    Dim T As clsTable
    Dim tG As clsTable

    Set T = NewCaseTable()

    Set tG = T.GroupBy("Dept", "Score:SUM")
    LogShim.AssertEq "gbsem_case_groups", 3, tG.RowCount ' 헤더 + IT/it 2그룹
    LogShim.AssertEq "gbsem_case_first_label", "IT", tG.Value(2, "Dept")
    LogShim.AssertEq "gbsem_case_second_label", "it", tG.Value(3, "Dept")
    LogShim.AssertEq "gbsem_case_first_sum", 30, tG.Value(2, "Score_SUM")
    LogShim.AssertEq "gbsem_case_second_sum", 5, tG.Value(3, "Score_SUM")

    ' 대조: 같은 원본에서 인덱스 경로는 대소문자 무시로 한 키 취급
    LogShim.AssertEq "gbsem_case_index_rows", 3, T.SelectRows("Dept", "it").Count
End Sub

' 현행: Empty 그룹 값은 CStr(Empty)="" 키로 묶임 (인덱스 계약의 Empty->""와 우연히 동치)
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

' 현행(문서화된 한계, R3 대상): 복합 키는 CStr & "|" 연결이라 구분자 포함 값이 충돌한다
' ("X|Y","Z")와 ("X","Y|Z")가 같은 키 "X|Y|Z|"로 병합됨
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
    LogShim.AssertEq "gbsem_collision_groups", 2, tG.RowCount ' 헤더 + 병합된 1그룹
    LogShim.AssertEq "gbsem_collision_sum", 30, tG.Value(2, "Score_SUM")
End Sub

' 현행: 그룹 컬럼 값은 키 문자열 재분해(Split)로 복원되므로 숫자도 String으로 나온다
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
