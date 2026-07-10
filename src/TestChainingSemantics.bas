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

' F3 (C1 계약): Compute의 미지 연산자/피연산자 컬럼 오타는 즉시 Err.Raise 9999,
' 타겟 컬럼 자동 생성 전에 검증되어 반쪽 실행 없음. C0 핀(조용한 무동작)에서 flip.
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

    ' (1) 미지 연산자 "%": raise + 타겟 컬럼 미생성 (반쪽 실행 없음)
    On Error Resume Next
    T.Compute "R1", "A", "%", "B"
    lErrNo1 = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "chain_f3_unknownop_errno", 9999, lErrNo1
    LogShim.AssertEq "chain_f3_unknownop_nocol", False, T.Headers.exists("R1")

    ' (2) 피연산자 컬럼 오타: raise + 타겟 컬럼 미생성
    On Error Resume Next
    T.Compute "R2", "A", "+", "NOPE"
    lErrNo2 = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "chain_f3_badcol_errno", 9999, lErrNo2
    LogShim.AssertEq "chain_f3_badcol_nocol", False, T.Headers.exists("R2")

    ' (3) 정상 경로 회귀: 유효 연산은 그대로 동작
    T.Compute "R3", "A", "+", "B"
    LogShim.AssertEq "chain_f3_valid_sum", 12, T.Value(2, "R3")
End Sub

' F4 (C1 계약): PatchIf 미지 연산자는 즉시 Err.Raise 9999 (modOperators 화이트리스트)
' C0 핀(조용한 무매칭)에서 flip. 데이터 수준 비교 오류(타입 불일치)는 무매칭 관용 유지.
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
    LogShim.AssertEq "chain_f4_patchif_errno", 9999, lErrNo
    LogShim.AssertEq "chain_f4_patchif_nocol", False, T.Headers.exists("STATUS")
End Sub

' C2 계약: canonical 종결자 ToTable(clsQuery/clsReportBuilder) + 즉시 실행 어휘 Keep
Public Sub Test_Chain_CanonicalAliases()
    Dim T As clsTable
    Dim tR As clsTable
    Dim vData(1 To 4, 1 To 2) As Variant

    vData(1, 1) = "Dept": vData(1, 2) = "Score"
    vData(2, 1) = "IT": vData(2, 2) = 10
    vData(3, 1) = "HR": vData(3, 2) = 20
    vData(4, 1) = "IT": vData(4, 2) = 30

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    ' clsQuery.ToTable == SelectTable
    Set tR = UnivDB.Query(T).Filter("Score", ">", 5).ToTable()
    LogShim.AssertEq "chain_c2_query_totable", 3, tR.RowCount - 1

    ' clsReportBuilder.Keep == Filter(즉시 실행), ToTable == Table
    Set tR = UnivDB.Report(T).Keep("Dept", "IS", "IT").ToTable()
    LogShim.AssertEq "chain_c2_report_keep_totable", 2, tR.RowCount - 1
End Sub

' C2 계약: SetVal(순수 setter, 빌더 반환)/GetVal(fail-fast getter) - Val 혼성 분리
Public Sub Test_Chain_BuilderSetGetVal()
    Dim B As clsTableBuilder
    Dim T As clsTable
    Dim v As Variant
    Dim lErrNo As Long

    Set B = UnivDB.Builder("ID")
    ' 체인 마지막 호출을 괄호 없이 쓰면 Call 불필요 (VBA 구문 계약 - 문서화는 C4)
    B.Row("P1").SetVal("Name", "Mouse").SetVal "Qty", 10

    LogShim.AssertEq "chain_c2_getval", 10, B.GetVal("Qty")

    ' GetVal은 미지 컬럼에 fail-fast (Val getter의 조용한 Empty와 대비)
    On Error Resume Next
    v = B.GetVal("NOPE")
    lErrNo = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "chain_c2_getval_badcol_errno", 9999, lErrNo

    Set T = B.ToTable()
    LogShim.AssertEq "chain_c2_setval_roundtrip", "Mouse", T.Value(2, "Name")
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
