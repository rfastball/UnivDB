# UnivDB 설계 메모: ACID 중 ACI — 어디까지 가고 어디서 멈추는가

작성일: 2026-07-10
성격: 탐색 메모 (결정 아님). 채택 시 설계 결정 메모로 승격 필요.
선행 문서: `UnivDB_design_decisions_202607.md` (rowID 세대 계약), `UnivDB_refactor_plan.md`

---

## 0. 문제 제기

UnivDB는 이름에 `DB`를 달고 있으나 트랜잭션·영속성·격리·쿼리 플래너·제약조건이 없다.
실체는 **Excel VBA를 위한 인메모리 DataFrame 라이브러리**다 (2D Variant 배열 + 헤더 Dictionary).

질문: 공수를 무시하면 ACID를 획득할 수 있는가?

판단 기준은 기존 문서와 동일하다.

```text
지향점: 배포용 라이브러리
  - "조용히 틀린 결과"가 최악의 실패 양식
규모: 1만 ~ 10만 행, Excel 배치 (적재 -> 변환 -> 보고)
```

결론 요약: **A·C·I는 fail-fast 철학의 완성이다. D와 MVCC는 다른 제품의 시작이다.**

---

## 1. MVCC는 채택하지 않는다

### 1.1 근거 1 — 독자가 0명이다

MVCC의 존재 이유는 "쓰기 중에도 읽는 쪽이 일관된 스냅샷을 본다"이다.
VBA는 단일 스레드이고, 매크로 실행 중 다른 코드가 같은 `clsTable`을 읽지 않는다.
동시 독자가 없는 곳에서 다중 버전을 유지하는 것은 순수 비용이다.

### 1.2 근거 2 — 연속 배열을 파괴한다

UnivDB의 성능 논지는 단 하나다.

> `.data`는 연속된 2D Variant 배열이며, `Range.Value = arr` 한 번으로 시트에 던진다.

행마다 버전 체인이 달리는 순간 `.data`와 `.RawArray`는 더 이상 그 배열이 아니다.
읽을 때마다 스냅샷을 materialize해야 하고, 셀 루프를 제거해 얻은 이득이 그 자리에서 증발한다.

**MVCC 도입 = 성능 논지 파기.** 검토 종료.

(참고: DuckDB의 MVCC가 "bulk-optimized"인 것은 OLAP의 거대 트랜잭션에 맞춰 튜닝했다는 뜻이지,
트랜잭션이 하나뿐인 세계를 위한 것이 아니다.)

---

## 2. A (원자성) — 가능. 재료가 이미 있다

### 2.1 설계 스케치: 스코프 가드

MVCC 대신 필요한 것은 undo log 또는 shadow copy다.
`clsReportBuilder`의 Copy-on-Write가 이미 shadow copy이며, 여기에 트랜잭션 스코프만 씌우면 된다.

VBA의 결정론적 참조 카운팅 소멸(`Class_Terminate`가 스코프 이탈 즉시 호출)을 이용하면
C++의 RAII 스코프 가드가 그대로 이식된다.

```vb
Dim tx As clsTx
Set tx = UnivDB.BeginTx("Emp", "Sales")   ' 대상 테이블 스냅샷 확보
UnivDB.Query("Emp").Where("Name", "Bob").SetVal("Dept", "Finance").Update
tx.Commit                                  ' Commit 없이 스코프 이탈 -> Terminate에서 자동 롤백
```

- `BeginTx`: 대상 테이블들의 `.data` / `.headers` / `Generation` 스냅샷 보관
- `Commit`: 스냅샷 폐기, 플래그 설정
- `Class_Terminate`: 플래그가 없으면 스냅샷 복원 (롤백)

### 2.2 구멍 1 — `End` 문과 디버거 중지

VBA의 `End` 문과 VBE 중지 버튼은 `Class_Terminate`를 **호출하지 않고** 상태를 통째로 파기한다.
언와인딩이 없으므로 스코프 가드가 뚫린다.

메모리 변경도 함께 증발하므로 실질 피해는 없다. **단, 다음 경우는 예외.**

### 2.3 구멍 2 (진짜 위험) — `DumpToSheet` 중 중단

시트 쓰기 도중 중단되면 시트가 반쯤 쓰인 채로 남는다.
메모리는 롤백되지만 시트는 롤백되지 않는다. 원자성 위반이 **관측 가능한 형태로** 남는 유일한 지점.

완화안 (택1 또는 병용):
- (a) 임시 시트에 전량 기록 후 `Move`/`Copy`로 교체 — 쓰기를 한 번의 연산으로 축소
- (b) `DumpToSheet` 진입 시 마커 셀 기록, 정상 종료 시 제거. `InitDB`가 잔존 마커를 발견하면 경고
- (c) 최소한 문서화: "`End` 문 사용 금지"를 코딩 컨벤션(10장)에 명문화

---

## 3. C (일관성) — 가장 남는 장사. 최우선 권장

### 3.1 현황

제약조건이 전무하다. 특히 `clsTableBuilder`는 존재하지 않는 컬럼을 조용히 생성한다.
이것은 "컬럼명 오타는 즉시 `Err 9999`"라는 6장의 계약과 정면으로 어긋난다.

### 3.2 제안: 스키마 제약을 6장 계약의 연장으로

6장 "Value & Key Semantics"가 **값 의미론**을 계약으로 확정했다면,
같은 형식으로 **스키마 의미론**을 확정한다.

| 제약 | 검사 시점 | 위반 시 |
|---|---|---|
| PK 유일성 | Insert / Upsert / Builder.ToTable | Err 9999 |
| NOT NULL | Insert / Update / PatchIf | Err 9999 |
| 타입 | Compute / Filter 피연산자 | Err 9999 (기존 규칙 확장) |
| FK (참조 테이블 존재) | JoinTable / FillFromMap | 선택적. 경고 또는 Err |

```vb
' 제안 API — 등록 시 계약 선언
UnivDB.RegisterTable "Inventory", Sheet1.Range("A1")
UnivDB.Constrain("Inventory").PrimaryKey("PROD_ID").NotNull("Qty")
```

### 3.3 근거

- fail-fast의 정확한 연장. `Err 9999`가 "컬럼명이 없다"에서 멈출 이유가 없다. "PROD_ID가 중복이다"까지 가야 한다.
- 공수가 평범하다. 배치 규모(1만~10만 행)에서 PK 유일성 검사는 기존 HashIndex 재사용으로 O(n).
- **ACID 네 글자 중 유일하게, 도입 비용이 아니라 도입 이득이 큰 항목.**

### 3.4 열린 질문

- `Builder`의 "없는 컬럼 자동 추가"는 편의 기능인가 버그인가? 제약 도입 시 이 동작은 opt-in으로 강등되어야 한다.
- 제약 위반이 배치 중간에 터지면 A(원자성)가 필요해진다. **C는 A를 요구한다.** 순서상 A가 먼저.

---

## 4. I (격리) — 트랜잭션 간에는 자명. 진짜 문제는 시트-메모리 간

### 4.1 트랜잭션 격리는 공짜다

동시 트랜잭션이 존재하지 않으므로 격리 수준을 논할 대상이 없다. Serializable이 자동 성립.

### 4.2 진짜 격리 대상: 시트

`RegisterTable`은 시트를 메모리로 복사한다. 그 순간부터 사용자는 스냅샷 격리 상태다.
그 사이에 다음 중 하나라도 일어나면:

- 사용자가 셀을 수정
- 다른 애드인이 동작
- 재계산 / 외부 데이터 새로고침

`DumpToSheet` 또는 시트 대상 쓰기가 그 변경을 **조용히 덮어쓴다.**
고전적 lost update이며, 이 라이브러리가 가장 싫어하는 실패 양식(조용히 틀림)이다.

### 4.3 제안: 낙관적 동시성 제어 (OCC)

MVCC로는 해결되지 않는다. 필요한 것은 원본 범위의 지문(fingerprint)이다.

```vb
' RegisterTable 시점
tbl.SourceFingerprint = HashRange(rng)   ' 값 + 크기 + 주소

' 시트 쓰기 직전
If HashRange(tbl.SourceRange) <> tbl.SourceFingerprint Then
    Err.Raise 9999, , "원본 시트가 등록 이후 변경되었습니다."
End If
```

### 4.4 기존 계약과의 구조적 동형

이것은 rowID 세대(Generation) 계약과 **정확히 같은 구조**다.

| | 파생 구조 | 원본 | 검증 |
|---|---|---|---|
| 기존 (메모리 내부) | HashIndex / RowSet / RowList | `clsTable` | `Generation` 불일치 -> Err 9999 |
| 제안 (메모리 ↔ 시트) | `clsTable` | 시트 Range | `Fingerprint` 불일치 -> Err 9999 |

즉 `clsTable` 자체가 시트의 파생 구조라는 사실을 인정하고, 이미 검증된 세대 계약을
한 층위 바깥으로 확장하는 것이다. 새 개념을 도입하지 않는다.

**개인 판단: ACID 전체보다 이 항목 하나가 실무 가치가 크다.**

---

## 5. D (지속성) — 여기서 멈춘다

### 5.1 기술적으로는 가능하다

`kernel32`의 `CreateFile` / `WriteFile` / `FlushFileBuffers`를 직접 선언해 WAL을 fsync하면 된다.
(VBA의 `Open` 문은 버퍼링 때문에 부적합.)

### 5.2 그러나 진실의 원본이 두 개가 된다

- WAL은 "커밋됨"이라고 말하는데 워크북은 저장되지 않았다
- 워크북 저장은 Excel이 수행하며, 우리 커밋과 원자적으로 묶이지 않는다

해결하는 유일한 방법은 **WAL을 진실의 원본으로 승격하고 시트를 뷰로 강등**한 뒤,
`InitDB`가 로그를 리플레이해 시트를 재구성하는 것이다.

그 순간 이 물건은 "Excel 시트를 편하게 다루는 라이브러리"가 아니라
**Excel을 렌더러로 쓰는 데이터베이스**가 된다.

### 5.3 판단

정체성 변경이다. 그리고 그 자리에는 이미 DuckDB가 있고 훨씬 잘한다.
**D는 채택하지 않는다.** 공수 문제가 아니라 범위 문제.

---

## 6. 종합

| 항목 | 가능? | 대가 | 권장 |
|---|---|---|---|
| **A** 원자성 | O | 스코프 가드. `End` 문 구멍은 봉인 불가 | 채택 (C의 전제) |
| **C** 일관성 | ◎ | 없음. 오히려 철학에 정합 | **최우선 채택** |
| **I** 격리 | O | 트랜잭션 격리는 자명. 시트-메모리 OCC가 실물 | **채택 (단독 가치 최대)** |
| **D** 지속성 | △ | 정체성 변경 | 기각 |
| MVCC | X | 연속 배열 파괴 = 성능 논지 파괴 | 기각 |

### 권장 순서

1. **I (OCC 지문)** — 단독으로 가치가 있고, 기존 Generation 계약의 확장이므로 개념 부채 없음
2. **A (clsTx 스코프 가드)** — C의 전제. Copy-on-Write 재사용
3. **C (제약조건 계층)** — 6장 계약을 스키마 의미론까지 확장

### 한 문장

> **ACI는 fail-fast의 완성이고, D는 다른 제품의 시작이다.**

---

## 7. 부수 관찰 (문서/명명)

> **후기 (2026-07-11):** 이 절이 제안한 카테고리 문구("인메모리 DataFrame 라이브러리")는
> 채택되지 않았다. 명명 논의는 `UnivDB_rename_proposal_202607.md`(판정: 전면 개명 보류, 경량
> 수정 채택)로 이어졌고, 채택된 자기소개 문구는 **"인메모리 관계 대수 엔진"**이다
> (`UnivDB_101.md` §1). 아래 관찰은 그 논의의 출발점으로서 원문 보존한다.

- `UnivDB_101.md` 1장의 자기소개("인메모리 DB 라이브러리")가 오해의 발원지.
  `DB`라는 기표는 트랜잭션·영속성·격리를 약속한다. 신입이 `Rollback`을 찾게 만든다.
- 개명 비용이 크므로 모듈명 `UnivDB`는 유지하되, 문서 첫 줄을 교체할 것을 제안:

  > **UnivDB — Excel VBA를 위한 인메모리 DataFrame 라이브러리.**
  > **게으르고 무뚝뚝합니다: 시키기 전엔 안 돌고, 틀리면 바로 소리칩니다.**

  앞줄이 카테고리(거짓말 안 함), 뒷줄이 철학(Query의 지연 실행 + fail-fast).
  뒷줄의 문장 구조는 R의 tibble이 검증한 것(`lazy and surly: do less and complain more`)에서 빌려옴.

- 선행 사례: `averrocchio/Pandas-Inspired-DataFrame-for-VBA` (VBA용 DataFrame, 2D Variant + 헤더 메타데이터).
  구조와 동기가 동일하나 Filter/Sort/Dedup/Join이 아직 계획 단계.
  다만 `Clean` / `InferTypes` / `Dedup` 동사는 UnivDB에 없음 — 검토 가치 있음.
- 카테고리 명명 참조: DFLib(Java, "DataFrame Library"), dplyr("a grammar of data manipulation"),
  DuckDB("in-process analytical database" — 얘는 ACID/MVCC가 실제로 있어서 `DB`를 붙일 자격이 있음).
