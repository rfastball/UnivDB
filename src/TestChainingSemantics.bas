Attribute VB_Name = "TestChainingSemantics"
' @desc 체이닝 API 조용한 실패 양식 회귀 고정 담당
' @depends UnivDB, LogShim
Option Explicit

' 체이닝 API 전수조사 (docs/tasks/2026-07-10_chaining_api_audit.md) Phase C0:
' S1 등급 위반(F1~F6)의 현행 실동작을 실행 테스트로 고정한다.
' 여기의 기대값은 "현행 실동작"이지 "옳은 계약"이 아니다 - C1(fail-fast화)에서
' 기대값 flip은 해당 fix와 같은 커밋으로 갱신된다. 전부 dict 경로만 사용.

' F1 (C1 계약): Filter 미지 연산자는 즉시 Err.Raise 9999 (modOperators 화이트리스트)
' C0 핀(에러 없이 0행)에서 flip - fail-fast fix와 같은 커밋
Public Sub Test_Chain_UnknownFilterOp()
    Dim T As clsTable
    Dim tR As clsTable
    Dim lErrNo As Long
    Dim vData(1 To 3, 1 To 1) As Variant

    vData(1, 1) = "Score"
    vData(2, 1) = 10
    vData(3, 1) = 30

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    On Error Resume Next
    Set tR = UnivDB.Query(T).Filter("Score", "=>", 10).SelectTable()
    lErrNo = Err.Number
    On Error GoTo 0

    LogShim.AssertEq "chain_f1_unknownop_errno", 9999, lErrNo
    LogShim.AssertEq "chain_f1_unknownop_nores", True, tR Is Nothing
End Sub

' F2 (C1 계약): GroupBy 미지 집계함수는 즉시 Err.Raise 9999 (modOperators 화이트리스트)
' C0 핀(에러 없이 빈 컬럼)에서 flip - fail-fast fix와 같은 커밋
Public Sub Test_Chain_UnknownAggFunc()
    Dim T As clsTable
    Dim tG As clsTable
    Dim lErrNo As Long
    Dim vData(1 To 3, 1 To 2) As Variant

    vData(1, 1) = "Dept": vData(1, 2) = "Score"
    vData(2, 1) = "IT": vData(2, 2) = 10
    vData(3, 1) = "IT": vData(3, 2) = 30

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    On Error Resume Next
    Set tG = T.GroupBy("Dept", "Score:SUMM")
    lErrNo = Err.Number
    On Error GoTo 0

    LogShim.AssertEq "chain_f2_unknownagg_errno", 9999, lErrNo
    LogShim.AssertEq "chain_f2_unknownagg_nores", True, tG Is Nothing

    ' 함수 생략 시 COUNT 기본은 유효한 계약으로 유지 (문서화는 C4)
    Set tG = T.GroupBy("Dept", "Score")
    LogShim.AssertEq "chain_f2_default_count", 2, tG.Value(2, "Score_COUNT")
End Sub

' F3: Compute는 미지 연산자도, 피연산자 컬럼 오타도 조용한 무동작
'     (타겟 컬럼은 자동 생성된 채 잔류 - 반쪽 실행)
Public Sub Test_Chain_ComputeSilentFailures()
    Dim T As clsTable
    Dim lErrNo1 As Long
    Dim lErrNo2 As Long
    Dim vData(1 To 3, 1 To 2) As Variant

    vData(1, 1) = "A": vData(1, 2) = "B"
    vData(2, 1) = 10: vData(2, 2) = 2
    vData(3, 1) = 30: vData(3, 2) = 3

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    ' (1) 미지 연산자 "%": 에러 없이 타겟 컬럼만 생성, 값 전부 Empty
    On Error Resume Next
    T.Compute "R1", "A", "%", "B"
    lErrNo1 = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "chain_f3_unknownop_errno", 0, lErrNo1
    LogShim.AssertEq "chain_f3_unknownop_created", True, T.Headers.exists("R1")
    LogShim.AssertEq "chain_f3_unknownop_empty", True, IsEmpty(T.Value(2, "R1"))

    ' (2) 피연산자 컬럼 오타: 에러 없이 Exit Sub, 타겟 컬럼만 생성
    On Error Resume Next
    T.Compute "R2", "A", "+", "NOPE"
    lErrNo2 = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "chain_f3_badcol_errno", 0, lErrNo2
    LogShim.AssertEq "chain_f3_badcol_created", True, T.Headers.exists("R2")
    LogShim.AssertEq "chain_f3_badcol_empty", True, IsEmpty(T.Value(2, "R2"))
End Sub

' F4: PatchIf 미지 연산자는 조용한 무매칭 (Case Else: bMatched=False)
Public Sub Test_Chain_PatchIfUnknownOp()
    Dim B As clsTableBuilder
    Dim T As clsTable
    Dim lErrNo As Long

    Set B = UnivDB.Builder("ID")
    Call B.Row("P1").Val("Qty", 0)

    On Error Resume Next
    B.PatchIf "Qty", "=>", 0, "Status", "SoldOut"
    lErrNo = Err.Number
    On Error GoTo 0

    Set T = B.ToTable()
    LogShim.AssertEq "chain_f4_patchif_errno", 0, lErrNo
    ' 무매칭이라 Status 컬럼 자체가 생성되지 않음
    LogShim.AssertEq "chain_f4_patchif_nocol", False, T.Headers.exists("STATUS")
End Sub

' F6 (C1 계약): 수치 연산자에 비숫자 비교값은 즉시 Err.Raise 9999 (문장 수준 오류)
' C0 핀(에러 없이 0행)에서 flip - fail-fast fix와 같은 커밋
Public Sub Test_Chain_NonNumericCompareValue()
    Dim T As clsTable
    Dim tR As clsTable
    Dim lErrNo As Long
    Dim vData(1 To 3, 1 To 1) As Variant

    vData(1, 1) = "Score"
    vData(2, 1) = 10
    vData(3, 1) = 30

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    On Error Resume Next
    Set tR = UnivDB.Query(T).Filter("Score", ">", "abc").SelectTable()
    lErrNo = Err.Number
    On Error GoTo 0

    LogShim.AssertEq "chain_f6_nonnumval_errno", 9999, lErrNo
    LogShim.AssertEq "chain_f6_nonnumval_nores", True, tR Is Nothing
End Sub
