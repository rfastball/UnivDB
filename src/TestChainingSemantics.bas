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

' C3 계약: clsTable 변형 메서드는 Me를 반환해 체인 유지 (audit F16, in-place 의미론 불변)
Public Sub Test_Chain_TableMutationChaining()
    Dim T As clsTable
    Dim tG As clsTable
    Dim vData(1 To 4, 1 To 2) As Variant

    vData(1, 1) = "Dept": vData(1, 2) = "Score"
    vData(2, 1) = "IT": vData(2, 2) = 30
    vData(3, 1) = "HR": vData(3, 2) = 20
    vData(4, 1) = "IT": vData(4, 2) = 10

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    ' GroupBy 결과에 곧바로 OrderBy 체인 (C3 이전엔 Sub라 컴파일 불가)
    Set tG = T.GroupBy("Dept", "Score:SUM").OrderBy("Score_SUM")
    LogShim.AssertEq "chain_c3_groupby_orderby_r2", 20, tG.Value(2, "Score_SUM")
    LogShim.AssertEq "chain_c3_groupby_orderby_r3", 40, tG.Value(3, "Score_SUM")

    ' AddColumn -> Compute -> OrderBy 변형 체인 (마지막 호출은 괄호 없이)
    T.AddColumn("Bonus", 5).Compute("Total", "Score", "+", "Bonus").OrderBy "Total", True
    LogShim.AssertEq "chain_c3_mutation_chain", 35, T.Value(2, "Total")
End Sub

' F23/F24 (audit 추보): CopyColumn 소스 컬럼 오타, Value Let 미지 컬럼 - fail-fast
' (이전: CopyColumn 조용한 무동작, Value Let은 헤드리스에서 안 보이는 Debug.Print 무시)
Public Sub Test_Chain_ColumnWriteFailFast()
    Dim T As clsTable
    Dim lErrNo1 As Long
    Dim lErrNo2 As Long
    Dim vData(1 To 2, 1 To 1) As Variant

    vData(1, 1) = "A"
    vData(2, 1) = 10

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    On Error Resume Next
    T.CopyColumn "NOPE", "X"
    lErrNo1 = Err.Number
    Err.Clear
    T.Value(2, "NOPE") = 1
    lErrNo2 = Err.Number
    On Error GoTo 0

    LogShim.AssertEq "chain_f23_copycol_errno", 9999, lErrNo1
    LogShim.AssertEq "chain_f23_copycol_nocol", False, T.Headers.exists("X")
    LogShim.AssertEq "chain_f24_valuelet_errno", 9999, lErrNo2
End Sub

' F17 (계약): clsQuery.Join - 조인을 쿼리 체인에 합류. 물질화 장벽 의미론.
' 기본 jType=jtInner(파사드 정합). vRight의 이름(String) 경로는 UnivDB.Join에
' verbatim 위임이라 기존 TestJoin(등록 경로 핀)이 커버 - 여기선 객체 경로만.
Public Sub Test_Chain_QueryJoinChain()
    Dim tL As clsTable
    Dim tR As clsTable
    Dim tRes As clsTable
    Dim vL(1 To 4, 1 To 2) As Variant
    Dim vR(1 To 4, 1 To 2) As Variant

    vL(1, 1) = "Dept": vL(1, 2) = "Amt"
    vL(2, 1) = "IT": vL(2, 2) = 10
    vL(3, 1) = "HR": vL(3, 2) = 20
    vL(4, 1) = "FIN": vL(4, 2) = 30   ' 우측 미매칭 - jtInner에서 탈락

    vR(1, 1) = "Dept": vR(1, 2) = "Boss"
    vR(2, 1) = "IT": vR(2, 2) = "Kim"
    vR(3, 1) = "IT": vR(3, 2) = "Park" ' IT 중복 키 (1:N 조인)
    vR(4, 1) = "HR": vR(4, 2) = "Lee"

    Set tL = UnivDB.CreateTable()
    tL.LoadFromArray vL
    Set tR = UnivDB.CreateTable()
    tR.LoadFromArray vR

    ' 기본 체인: 필터 -> 조인(jtInner 기본) -> prefix 컬럼 위 필터 -> 종결
    Set tRes = UnivDB.Query(tL).Filter("Amt", ">", 5) _
        .Join(tR, "Dept", , "R_") _
        .Filter("R_Boss", "IS", "Kim").ToTable()
    LogShim.AssertEq "chain_f17_join_rows", 1, tRes.RowCount - 1
    LogShim.AssertEq "chain_f17_join_prefixcol", True, tRes.Headers.exists("R_BOSS")
    LogShim.AssertEq "chain_f17_join_amt", 10, tRes.Value(2, "Amt")

    ' jtInner 기본 확인: 필터 없이 조인만 - FIN(미매칭) 탈락, IT는 1:N으로 2행
    Set tRes = UnivDB.Query(tL).Join(tR, "Dept").ToTable()
    LogShim.AssertEq "chain_f17_join_inner_default", 3, tRes.RowCount - 1
End Sub

' F17 (계약): 물질화 장벽 - Join 전 정렬/Top은 좌측에 소진되고, Join 후
' 파이프라인 상태(투영/정렬/Top)는 초기값으로 리셋된다.
Public Sub Test_Chain_QueryJoinBarrier()
    Dim tL As clsTable
    Dim tR As clsTable
    Dim tRes As clsTable
    Dim vL(1 To 4, 1 To 2) As Variant
    Dim vR(1 To 4, 1 To 2) As Variant

    vL(1, 1) = "Dept": vL(1, 2) = "Amt"
    vL(2, 1) = "IT": vL(2, 2) = 10
    vL(3, 1) = "HR": vL(3, 2) = 20
    vL(4, 1) = "FIN": vL(4, 2) = 30

    vR(1, 1) = "Dept": vR(1, 2) = "Boss"
    vR(2, 1) = "IT": vR(2, 2) = "Kim"
    vR(3, 1) = "IT": vR(3, 2) = "Park"
    vR(4, 1) = "HR": vR(4, 2) = "Lee"

    Set tL = UnivDB.CreateTable()
    tL.LoadFromArray vL
    Set tR = UnivDB.CreateTable()
    tR.LoadFromArray vR

    ' 장벽: OrderBy+Top(1)이 좌측에 소진 - Amt 최대(FIN 30) 1행만 좌측, jtLeft로 보존
    Set tRes = UnivDB.Query(tL).OrderBy("Amt", True).Top(1) _
        .Join(tR, "Dept", jtLeft, "R_").ToTable()
    LogShim.AssertEq "chain_f17_barrier_rows", 1, tRes.RowCount - 1
    LogShim.AssertEq "chain_f17_barrier_topval", 30, tRes.Value(2, "Amt")

    ' 리셋: Join 전 Top(1)이 좌측(IT 1행)에 소진된 뒤, 조인(1:N)의 2행이
    ' 이전 Top에 잘리지 않고 전부 반환 - m_lLimit 잔존 없음 증명
    Set tRes = UnivDB.Query(tL).Top(1).Join(tR, "Dept", jtInner, "R_").ToTable()
    LogShim.AssertEq "chain_f17_reset_rows", 2, tRes.RowCount - 1
End Sub

' F25 (audit 추보): 조인 키 컬럼 오타는 즉시 Err.Raise 9999.
' 이전 동작(MsgBox 모달 + Nothing 반환)은 헤드리스에서 핀 고정 자체가 불가능해
' 2단 커밋 패턴 예외 - fix와 신규 계약 핀만 (사유는 audit 문서 기록).
Public Sub Test_Chain_JoinBadKeyFailFast()
    Dim tL As clsTable
    Dim tR As clsTable
    Dim tJ As clsTable
    Dim lErrNo1 As Long
    Dim lErrNo2 As Long
    Dim vL(1 To 2, 1 To 1) As Variant
    Dim vR(1 To 2, 1 To 1) As Variant

    vL(1, 1) = "K": vL(2, 1) = "a"
    vR(1, 1) = "K": vR(2, 1) = "a"

    Set tL = UnivDB.CreateTable()
    tL.LoadFromArray vL
    Set tR = UnivDB.CreateTable()
    tR.LoadFromArray vR

    On Error Resume Next
    Set tJ = UnivDB.Join(tL, tR, "NOPE")
    lErrNo1 = Err.Number
    Err.Clear
    Set tJ = UnivDB.Join(tL, tR, "K", jtInner, "", "NOPE")
    lErrNo2 = Err.Number
    On Error GoTo 0

    LogShim.AssertEq "chain_f25_join_badleftkey_errno", 9999, lErrNo1
    LogShim.AssertEq "chain_f25_join_badrightkey_errno", 9999, lErrNo2
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
