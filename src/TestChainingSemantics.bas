Attribute VB_Name = "TestChainingSemantics"
' @desc 체이닝 API 조용한 실패 양식 회귀 고정 담당
' @depends UnivDB, LogShim
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
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

' B축 (계약): 값 스테이징/조건부 패치 어휘 수렴 - clsQuery.SetVal(=SetValue),
' RB.PatchIf(=UpdateValue, 빌더 PatchIf 인자 순서 cond,op,val,target,new)
Public Sub Test_Chain_VocabAliases()
    Dim T As clsTable
    Dim tRes As clsTable
    Dim lChanged As Long
    Dim vData(1 To 3, 1 To 2) As Variant

    vData(1, 1) = "Name": vData(1, 2) = "Grade"
    vData(2, 1) = "A": vData(2, 2) = 100
    vData(3, 1) = "B": vData(3, 2) = 50

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    ' clsQuery.SetVal == SetValue: 스테이징 후 Update
    lChanged = UnivDB.Query(T).Filter("Grade", ">=", 100).SetVal("Name", "TOP").Update()
    LogShim.AssertEq "chain_b_setval_changed", 1, lChanged
    LogShim.AssertEq "chain_b_setval_applied", "TOP", T.Value(2, "Name")

    ' RB.PatchIf(cond, op, condVal, target, new) == UpdateValue(target, new, cond, op, condVal)
    Set tRes = UnivDB.Report(T).PatchIf("Grade", "<", 60, "Name", "LOW").ToTable()
    LogShim.AssertEq "chain_b_patchif_applied", "LOW", tRes.Value(3, "Name")
    LogShim.AssertEq "chain_b_patchif_untouched", "TOP", tRes.Value(2, "Name")
End Sub

' B축 (계약): clsTable.AddColumns/RemoveColumns - 개별 인자/csv 허용(SelectColumns
' 파싱 규약 공유), AddColumn/RemoveColumn 멱등 계약 위임, 체인 유지
Public Sub Test_Chain_TableAddRemoveColumns()
    Dim T As clsTable
    Dim vData(1 To 2, 1 To 1) As Variant

    vData(1, 1) = "A"
    vData(2, 1) = 1

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    ' csv 한 줄 + 체인
    T.AddColumns("B, C").AddColumns "D"
    LogShim.AssertEq "chain_b_addcols_csv", True, T.Headers.exists("C")
    LogShim.AssertEq "chain_b_addcols_chain", True, T.Headers.exists("D")
    LogShim.AssertEq "chain_b_addcols_count", 4, T.ColCount

    ' 개별 인자 + 멱등 ensure (기존 B는 무시)
    T.AddColumns "E", "B"
    LogShim.AssertEq "chain_b_addcols_multi", 5, T.ColCount

    ' 제거: csv + 체인 + 멱등 drop (없는 X 무시)
    T.RemoveColumns("D, E").RemoveColumns "X"
    LogShim.AssertEq "chain_b_removecols", 3, T.ColCount
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

' ==============================================================================
' C-1: FilterAny (교차 컬럼 OR) - 신규 계약 핀 (impl과 같은 커밋, 1ce5f27 패턴)
' ==============================================================================

' FilterAny 공용 테이블: Dept/Amt/Active 3컬럼 5행
Private Function NewOrGroupTable() As clsTable
    Dim T As clsTable
    Dim vData(1 To 6, 1 To 3) As Variant

    vData(1, 1) = "Dept": vData(1, 2) = "Amt": vData(1, 3) = "Active"
    vData(2, 1) = "Sales": vData(2, 2) = 500: vData(2, 3) = "Y"
    vData(3, 1) = "IT": vData(3, 2) = 1500: vData(3, 3) = "Y"
    vData(4, 1) = "HR": vData(4, 2) = 800: vData(4, 3) = "N"
    vData(5, 1) = "Sales": vData(5, 2) = 2000: vData(5, 3) = "N"
    vData(6, 1) = "IT": vData(6, 2) = 300: vData(6, 3) = "Y"

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData
    Set NewOrGroupTable = T
End Function

' 기본 계약: (arm1 OR arm2) AND 후속 Filter
Public Sub Test_Chain_FilterAnyCrossColumnOr()
    Dim T As clsTable
    Dim tR As clsTable

    Set T = NewOrGroupTable()

    ' (Dept=Sales {2,5} OR Amt>1000 {3,5}) = {2,3,5}, AND Active=Y = {2,3}
    Set tR = UnivDB.Query(T) _
        .FilterAny(Array("Dept", "IS", "Sales"), Array("Amt", ">", 1000)) _
        .Filter("Active", "IS", "Y") _
        .ToTable()

    LogShim.AssertEq "chain_c1_or_and_rows", 3, tR.RowCount
    LogShim.AssertEq "chain_c1_or_and_r1", "Sales", tR.Value(2, "Dept")
    LogShim.AssertEq "chain_c1_or_and_r2", "IT", tR.Value(3, "Dept")
End Sub

' 첫 조건 시딩(전행) / 선행 Where 후보 축소와의 결합
Public Sub Test_Chain_FilterAnyFirstCondition()
    Dim T As clsTable
    Dim tR As clsTable

    Set T = NewOrGroupTable()

    ' 첫 조건: 전행 시딩 후 (HR {4} OR Amt<400 {6}) = {4,6}
    Set tR = UnivDB.Query(T) _
        .FilterAny(Array("Dept", "IS", "HR"), Array("Amt", "<", 400)) _
        .ToTable()
    LogShim.AssertEq "chain_c1_first_rows", 3, tR.RowCount

    ' 선행 Where(Active=Y {2,3,6}) 후: (Sales{2} OR Amt>1000{3}) = {2,3}
    Set tR = UnivDB.Query(T) _
        .Where("Active", "Y") _
        .FilterAny(Array("Dept", "IS", "Sales"), Array("Amt", ">", 1000)) _
        .ToTable()
    LogShim.AssertEq "chain_c1_after_where_rows", 3, tR.RowCount
End Sub

' 두 arm에 모두 매칭되는 행은 결과에 1회만 (비트맵 멱등)
Public Sub Test_Chain_FilterAnyOverlapNoDupes()
    Dim T As clsTable
    Dim tR As clsTable

    Set T = NewOrGroupTable()

    ' Sales {2,5} OR Amt>1800 {5} - 행5는 양쪽 매칭이지만 1회
    Set tR = UnivDB.Query(T) _
        .FilterAny(Array("Dept", "IS", "Sales"), Array("Amt", ">", 1800)) _
        .ToTable()
    LogShim.AssertEq "chain_c1_overlap_rows", 3, tR.RowCount
End Sub

' 동일 컬럼 arm = IN 의미론 (값 목록 매칭)
Public Sub Test_Chain_FilterAnySameColumnIn()
    Dim T As clsTable
    Dim tR As clsTable

    Set T = NewOrGroupTable()

    Set tR = UnivDB.Query(T) _
        .FilterAny(Array("Dept", "IS", "IT"), Array("Dept", "IS", "HR")) _
        .ToTable()
    LogShim.AssertEq "chain_c1_in_rows", 4, tR.RowCount
End Sub

' 동적 arm 목록 (arm 배열 1개) = ParamArray 개별 인자와 동일 결과
Public Sub Test_Chain_FilterAnyDynamicList()
    Dim T As clsTable
    Dim tR1 As clsTable
    Dim tR2 As clsTable
    Dim vArmList As Variant

    Set T = NewOrGroupTable()
    vArmList = Array(Array("Dept", "IS", "IT"), Array("Dept", "IS", "HR"))

    Set tR1 = UnivDB.Query(T).FilterAny(vArmList).ToTable()
    Set tR2 = UnivDB.Query(T).FilterAny(Array("Dept", "IS", "IT"), Array("Dept", "IS", "HR")).ToTable()
    LogShim.AssertEq "chain_c1_dynlist_rows", tR2.RowCount, tR1.RowCount
End Sub

' 문장 수준 검증 5종: 0 arm / 2요소 arm / 미지 연산자 / 컬럼 오타 / 비숫자 비교값
Public Sub Test_Chain_FilterAnyValidation()
    Dim T As clsTable
    Dim lErrNo As Long

    Set T = NewOrGroupTable()

    On Error Resume Next
    UnivDB.Query(T).FilterAny
    lErrNo = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "chain_c1_val_zeroarm", 9999, lErrNo

    On Error Resume Next
    UnivDB.Query(T).FilterAny Array("Dept", "IS")
    lErrNo = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "chain_c1_val_2elem", 9999, lErrNo

    On Error Resume Next
    UnivDB.Query(T).FilterAny Array("Dept", "=>", "Sales")
    lErrNo = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "chain_c1_val_unknownop", 9999, lErrNo

    On Error Resume Next
    UnivDB.Query(T).FilterAny Array("Nope", "IS", "x")
    lErrNo = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "chain_c1_val_badcol", 9999, lErrNo

    On Error Resume Next
    UnivDB.Query(T).FilterAny Array("Amt", ">", "abc")
    lErrNo = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "chain_c1_val_nonnum", 9999, lErrNo
End Sub

' 검증은 변형 앞: arm2 오류 시 arm1의 부분 축소 없음 (audit F3 교훈)
Public Sub Test_Chain_FilterAnyValidateBeforeMutate()
    Dim T As clsTable
    Dim q As clsQuery
    Dim lErrNo As Long

    Set T = NewOrGroupTable()
    Set q = UnivDB.Query(T).Where("Active", "Y") ' 후보 {2,3,6}

    On Error Resume Next
    q.FilterAny Array("Dept", "IS", "Sales"), Array("Nope", "IS", "x")
    lErrNo = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "chain_c1_prevalidate_errno", 9999, lErrNo

    ' 후보 무변화 확인: Where 결과 3행 그대로
    LogShim.AssertEq "chain_c1_prevalidate_intact", 4, q.ToTable().RowCount
End Sub

' clsReportBuilder.KeepAny: 즉시 실행 OR + 원본 무접촉
Public Sub Test_Chain_ReportKeepAny()
    Dim T As clsTable
    Dim tR As clsTable

    Set T = NewOrGroupTable()

    Set tR = UnivDB.Report(T) _
        .KeepAny(Array("Dept", "IS", "Sales"), Array("Amt", ">", 1000)) _
        .ToTable()

    LogShim.AssertEq "chain_c1_keepany_rows", 4, tR.RowCount
    LogShim.AssertEq "chain_c1_keepany_orig_intact", 6, T.RowCount
End Sub

' ==============================================================================
' C-2: Compute 숫자 상수 피연산자 - 신규 계약 핀 (impl과 같은 커밋)
' ==============================================================================

' 기본 계약: 컬럼 x 상수 / 상수 x 컬럼 / 체인 유지
Public Sub Test_Chain_ComputeConstOperand()
    Dim T As clsTable
    Dim tChain As clsTable

    Set T = NewOrGroupTable()

    T.Compute "Half", "Amt", "*", 0.5
    LogShim.AssertEq "chain_c2_const_right", 250, T.Value(2, "Half")

    T.Compute "Rem", 100, "-", "Amt"
    LogShim.AssertEq "chain_c2_const_left", -400, T.Value(2, "Rem")

    Set tChain = T.Compute("Twice", "Amt", "*", 2)
    LogShim.AssertEq "chain_c2_const_chain", True, tChain Is T
End Sub

' 상수 x 상수: 전 행 동일 값 (금지보다 단순한 계약 - AddColumn 초기화와 중복일 뿐)
Public Sub Test_Chain_ComputeConstBoth()
    Dim T As clsTable

    Set T = NewOrGroupTable()
    T.Compute "K", 2, "*", 3
    LogShim.AssertEq "chain_c2_both_r1", 6, T.Value(2, "K")
    LogShim.AssertEq "chain_c2_both_r5", 6, T.Value(6, "K")
End Sub

' 문장 수준 검증: String "3"은 여전히 컬럼명(9999) / Empty 거부 / 상수 0 나눗셈 거부
' 전부 타겟 컬럼 미생성 (반쪽 실행 없음 - audit F3 계약 유지)
Public Sub Test_Chain_ComputeConstValidation()
    Dim T As clsTable
    Dim lErrNo As Long

    Set T = NewOrGroupTable()

    On Error Resume Next
    T.Compute "R1", "Amt", "+", "3"
    lErrNo = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "chain_c2_strlit_errno", 9999, lErrNo
    LogShim.AssertEq "chain_c2_strlit_nocol", False, T.Headers.exists("R1")

    On Error Resume Next
    T.Compute "R2", "Amt", "+", Empty
    lErrNo = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "chain_c2_empty_errno", 9999, lErrNo
    LogShim.AssertEq "chain_c2_empty_nocol", False, T.Headers.exists("R2")

    On Error Resume Next
    T.Compute "R3", "Amt", "/", 0
    lErrNo = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "chain_c2_divzero_errno", 9999, lErrNo
    LogShim.AssertEq "chain_c2_divzero_nocol", False, T.Headers.exists("R3")
End Sub

' 상수와 & 결합은 CStr 참여, 비숫자 셀은 데이터 수준 skip 관용 유지
Public Sub Test_Chain_ComputeConstConcatAndNullSkip()
    Dim T As clsTable
    Dim vData(1 To 4, 1 To 1) As Variant

    Set T = NewOrGroupTable()
    T.Compute "Lbl", "Dept", "&", 9
    LogShim.AssertEq "chain_c2_concat", "Sales9", T.Value(2, "Lbl")

    vData(1, 1) = "A"
    vData(2, 1) = 10
    vData(3, 1) = "abc"
    vData(4, 1) = 30

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData
    T.Compute "R", "A", "+", 1
    LogShim.AssertEq "chain_c2_skip_ok1", 11, T.Value(2, "R")
    LogShim.AssertEq "chain_c2_skip_cell", True, IsEmpty(T.Value(3, "R"))
    LogShim.AssertEq "chain_c2_skip_ok2", 31, T.Value(4, "R")
End Sub

' 정렬 키가 투영에 없어도 정렬/Top 의미를 보존하며 원본은 바뀌지 않는다.
Public Sub Test_Chain_SortProjectionAndEmpty()
    Dim T As clsTable, result As clsTable, gen As Long
    Set T = NewOrGroupTable()
    gen = T.Generation
    Set result = UnivDB.Query(T).Where("Active", "Y").OrderBy("Amt", True).SelectColumns("Dept").Top(1).ToTable()
    LogShim.AssertEq "sort_projection_top", "IT", result.Value(2, "Dept")
    LogShim.AssertEq "sort_projection_cols", 1, result.ColCount
    LogShim.AssertEq "sort_projection_rows", 2, result.RowCount
    LogShim.AssertEq "sort_source_generation", gen, T.Generation
    LogShim.AssertEq "sort_source_first", 500, T.Value(2, "Amt")

    Set result = UnivDB.Query(T).Where("Dept", "__none__").OrderBy("Amt").Top(1).ToTable()
    LogShim.AssertEq "sort_empty_rows", 1, result.RowCount
    result.OrderBy "Amt", True
    LogShim.AssertEq "sort_empty_direct", 1, result.RowCount
End Sub

' 문장 오류는 후보 0행에서도 동일하게 실패한다.
Public Sub Test_Chain_ValidationBeforeEmpty()
    Dim T As clsTable, q As clsQuery, result As clsTable
    Dim emptyFirst As Long, n As Long
    Set T = NewOrGroupTable()
    For emptyFirst = 0 To 1
        Set q = UnivDB.Query(T)
        If emptyFirst = 1 Then q.Where "Dept", "__none__"
        On Error Resume Next
        Err.Clear
        q.Where "__bad__", "IT"
        n = Err.Number
        On Error GoTo 0
        LogShim.AssertEq "where_bad_column", 9999, n

        On Error Resume Next
        Err.Clear
        q.Filter "__bad__", "IS", "IT"
        n = Err.Number
        On Error GoTo 0
        LogShim.AssertEq "filter_bad_column", 9999, n

        On Error Resume Next
        Err.Clear
        q.Filter "Amt", ">", "abc"
        n = Err.Number
        On Error GoTo 0
        LogShim.AssertEq "filter_bad_value", 9999, n

        On Error Resume Next
        Err.Clear
        Set result = q.OrderBy("__bad__").ToTable()
        n = Err.Number
        On Error GoTo 0
        LogShim.AssertEq "sort_bad_column", 9999, n
    Next emptyFirst
End Sub

' 모든 SET 컬럼을 쓰기 전에 검증한다.
Public Sub Test_Chain_UpdateValidatesAllColumns()
    Dim T As clsTable, q As clsQuery, n As Long, emptyFirst As Long
    For emptyFirst = 0 To 1
        Set T = NewOrGroupTable()
        Set q = UnivDB.Query(T)
        If emptyFirst = 1 Then q.Where "Dept", "__none__"
        q.SetVal("Amt", 111).SetVal "__bad__", 1
        On Error Resume Next
        Err.Clear
        q.Update
        n = Err.Number
        On Error GoTo 0
        LogShim.AssertEq "update_bad_column", 9999, n
        LogShim.AssertEq "update_no_partial_write", 500, T.Value(2, "Amt")
    Next emptyFirst
End Sub
