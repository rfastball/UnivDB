# UnivDB 설계 결정 메모 v2: rowID 정체성과 스토리지 레이아웃

작성일: 2026-07-05
선행 문서: `UnivDB_refactor_plan.md` (내부 실행 표현 전환, Phase 0~8)
후속 반영: `tasks/2026-07-05_next_direction.md` (실행 계획)

---

## 0. 문서 목적과 판단 기준

설계 결정 메모 v1(9항)은 내부 실행 언어(HashIndex / RowList / RowSet), Dictionary 밀봉, Core 커널 경계를 결정했다. 이 문서는 그 위 층위의 두 결정을 확정한다.

- 결정 A: rowID의 정체성 의미론
- 결정 C: 스토리지 레이아웃

모든 판단의 기준은 다음 두 가지다.

```text
지향점: 배포용 라이브러리
  - API 안정성, 명시적 계약, fail-fast가 우선
  - "조용히 틀린 결과"가 최악의 실패 양식

규모: 1만 ~ 10만 행
  - Excel 배치 워크로드 (적재 -> 변환 -> 보고)
  - 구조 변경은 배치 단위로 발생, 행 단위 상호 변경은 드묾
```

---

## 1. 결정 A: rowID는 물리 행 번호 + 세대(Generation) 계약이다

### 1.1 결정

rowID는 안정적 논리 식별자(surrogate ID)로 승격하지 않는다. rowID의 정의는 다음과 같다.

> rowID는 특정 테이블의 특정 세대(Generation) 안에서만 유효한 물리 행 번호다.

- 모든 구조 변경(Insert / Delete / OrderBy / Upsert / 컬럼 추가·삭제)은 `Table.Generation`을 증가시킨다.
- 모든 파생 구조(HashIndex, RowSet, RowList)는 생성 시점의 세대를 기억한다.
- 세대가 불일치한 파생 구조를 사용하면 Err.Raise로 즉시 실패한다 (fail-fast).

### 1.2 근거

surrogate ID를 배제하는 이유:

```text
- 1만~10만 행 배치 워크로드에서는 surrogate ID의 이득
  (증분 인덱스 유지, 교차 변경 참조)이 실현되지 않는다.
- VBA에서 배치 변경 후 인덱스 재빌드가 per-row 증분 유지보다 오히려 빠르다.
- 물리 rowID는 m_vData(rowId, col) 직접 접근으로 최저 비용이다.
  간접화 계층은 모든 셀 접근에 세금을 얹는다.
```

세대 계약을 도입하는 이유:

```text
- 현재도 "구조 변경 후 이전 파생물 사용 금지"는 암묵적 계약이다.
- 배포용 라이브러리는 이 계약을 위반한 사용자를 조용히 틀린 결과가 아니라
  명시적 에러로 막아야 한다.
- 세대 번호 비교는 Long 비교 1회 - 비용이 사실상 0이다.
```

### 1.3 계약이 정의하는 것

이 결정으로 다음 의미론이 자동으로 정해진다.

```text
- Delete / OrderBy 후 이전 쿼리 결과(RowSet) 재사용: 에러 (이전에는 미정의 동작)
- 물질화된 결과(SelectTable)는 새 테이블: 테이블 간 행 정체성 없음
- 인덱스 무효화: 세대 규칙의 특수 사례로 흡수 (아래 3장)
```

---

## 2. 결정 C: 정본 스토리지는 row-major 2D Variant, 접근은 컬럼 표면 뒤에 밀봉

### 2.1 결정

- 정본(canonical) 스토리지는 row-major 2D Variant 배열을 유지한다.
- 단, 실행계(Where / Filter / Join / GroupBy / 인덱스 빌드)는 `m_vData`를 직접 만지지 않고 컬럼 접근 표면(ColumnReader 또는 동등)을 통해서만 접근한다.
- columnar / typed 전환은 전면 도입하지 않는다. 병목이 측정된 컬럼만 1D 배열로 지연 물질화하는 **핫 컬럼 캐시**로 선택 도입한다.

### 2.2 근거

전면 columnar를 배제하는 이유:

```text
- 데이터의 입구와 출구가 모두 Range.Value2 = 2D Variant다.
- 전면 columnar는 적재와 출력 양방향에서 O(N x C) 전치 비용을 낸다.
  (10만 행 x 20열 = 200만 Variant 복사 왕복)
- 1만~10만 행에서 2D Variant 스캔은 병목이되 치명적이지 않다.
```

컬럼 접근 밀봉을 택하는 이유:

```text
- 진짜 핫패스(필터 스캔, 인덱스 빌드)는 단일 컬럼 순회다.
  해당 컬럼만 1D 캐시로 물질화하면 국소적으로 해결된다.
- Dictionary를 DictHashIndex에 밀봉한 것과 동형의 조치다:
  레이아웃이 표면 뒤로 들어가면, 향후 columnar 전환도
  전면 재작성이 아니라 백엔드 교체 문제로 축소된다.
- 핫 컬럼 캐시는 CoreHashIndex보다 큰 성능 지렛대일 수 있다.
  Phase 7 벤치에서 두 지렛대를 비교 측정한다.
```

---

## 3. 통합 원리: 모든 파생 구조는 (테이블, 세대)에 결속된다

결정 A와 C에서 하나의 상위 원리가 도출된다.

> HashIndex, RowSet, RowList, 컬럼 캐시, (미래의) 정렬 캐시는 전부
> (테이블, 세대) 쌍에 결속된 파생물이며, 하나의 무효화 규칙을 공유한다.

- 구조 변경 = 세대 증가 = 모든 파생물 실효. 개별 무효화 정책을 파생물마다 설계하지 않는다.
- refactor plan Phase 1의 인덱스 무효화 정비는 이 원리의 특수 사례였다. 규칙을 일반화한다.
- 새 파생 구조를 추가할 때의 질문은 하나로 줄어든다: "세대를 기억하는가, 불일치 시 실패하는가."

---

## 4. 명시적 비결정 (이월)

다음 축은 이번에 결정하지 않는다.

```text
B. 실행 모델 (즉시 평가 유지 vs 논리 플랜/물리 실행 분리)
D. NULL / 타입 의미론 (Empty=NULL? 3치 논리? 컬럼 타입 선언?)
E. 변경 모델 (in-place + 세대 vs 스냅샷 기반)
F. UnivCore와의 장기 관계 (독립 커널 vs 흡수)
```

단, 하나는 기한이 있다.

> D(NULL/타입 의미론)는 배포용 라이브러리 API 계약의 일부다.
> 1.0 공개 전에 반드시 확정한다. GroupBy의 대소문자 의미론 불일치
> (binary compare vs 인덱스의 vbTextCompare)도 이 결정의 하위 항목이다.

---

## 5. 로드맵 영향 요약

- 세대 계약 도입이 Phase 6.5(레거시 인덱스 경로 단일화)보다 선행하는 축이 된다.
- 컬럼 접근 밀봉은 Phase 6.7(DbHashIndex 인터페이스)과 같은 층위의 표면 정리 작업이다.
- Phase 7 준비 벤치에 "핫 컬럼 캐시 vs CoreHashIndex" 지렛대 비교를 추가한다.
- 상세 실행 계획은 `tasks/2026-07-05_next_direction.md`에 반영한다.

---

## 6. 한 문장 요약

> rowID는 승격하지 않고 계약을 승격한다.
> 레이아웃은 바꾸지 않고 접근을 밀봉한다.
> 무효화는 파생물마다 설계하지 않고 세대 하나로 다스린다.

---

## 부록 (2026-07-11 추가): fail-fast의 구조적 기원

`UnivDB_rename_proposal_202607.md` §2.3의 통찰을 설계 결정으로 승격한다 (개명 제안 자체는
보류됐으나, 이 인과는 이름과 독립이므로 별도 보존한다).

이 라이브러리 제1원칙("조용히 틀린 결과가 최악의 실패 양식")이 왜 존재하는가에 대한 구조적 답:

```text
- 집합 연산의 고유한 실패 양식은 "빈 집합"이다.
    - 교집합이 비어 0행  -> 유효한 결과
    - 컬럼명 오타로 0행   -> 버그
    - 두 경우가 반환값에서 구별되지 않는다.
- 열(column)을 지시하는 DataFrame 계열은 오타가 즉시 KeyError로 드러난다.
  집합을 지시하는 순간, 오타가 "유효한 답"으로 위장한다.
```

따라서 `Err 9999`(fail-fast)는 나중에 덧붙인 철학이 아니라, 집합 연산을 바닥에 둔 라이브러리가
반드시 마주치는 구조적 문제에 대한 **필연적 응답**이다. 이 계보(필터 -> 집합 연산 -> 체이닝)는
내부 실행 표현이 `RowSet`/`clsSet.Intersection`을 쓰는 사실로도 코드에 남아 있다.

## 부록 (2026-07-13 추가): 결정 E - 인덱스 페이로드는 CSR 평탄 배열이다

### E.1 결정

해시 인덱스(DictHashIndex/CoreHashIndex)의 rowId 페이로드는 키당 `DbRowList` 객체가
아니라 공용 `DbCsrPayload` 하나가 보관한다: 공유 Long 배열(mRowIds) + 키별
(offset, count). 백엔드는 키->keyId(1기반 밀집) 매핑만 담당한다
(dict = Count+1 순차 부여, core = FastHash slot).

### E.2 근거

키당 객체 저장은 고유키성 컬럼(고유키 수 ~= 행 수)에서 수만 개 COM 인스턴스를 만들고,
ClearDB/Terminate 시 개별 소멸 캐스케이드가 teardown 지배원이 된다
(20k행 실측 60ms, 데이터 배열 해제는 2ms - `2026-07-13_teardown_cost_index_payload.md`).
CSR 전환 후 소멸은 Long 배열 Erase 몇 회로 키 수와 무관해진다 (실측 1ms).

### E.3 규약

- **빌드**: AddRow 스트리밍이 엔트리 버퍼(keyId, rowId 쌍)에 누적, 첫 읽기 접근자에서
  lazy finalize(prefix-sum + 안정 counting-sort scatter). 키 내 rowId 순서 = AddRow
  호출 순서. AddRow는 dirty 마킹만 하므로 읽기 후 재적재도 안전(재finalize).
  빌드 버퍼는 finalize 후에도 유지된다(재finalize 원본, 행당 Long 2개).
- **복사 반환 의미론**: `TryGetRows`/`DbIndexCursor.Rows`는 인덱스 내부의 라이브 참조가
  아니라 세대 결속(BindRaw)을 탑재한 fresh DbRowList 복사본을 반환한다. 반환 후의
  인덱스 변경은 복사본에 반영되지 않는다.
- **span 버퍼 소유권**: `TryGetRowIdSpan`/`TryGetRowCount`/`TryGetFirstRowId`(인터페이스),
  커서 `RowCount`/`CopyRows`는 객체 생성 없는 핫 경로용. 버퍼는 호출자 소유 재사용,
  구현이 부족 시 배가 ReDim, 내용은 호출마다 덮어씀 - 소비는 outCount 기준.
  miss/Error 키는 False + outCount 0.
- **세대 결속**: 인덱스 단위 결속(mOwnerPtr/mGeneration)이 유일 진실이고 복사본에
  전파된다. 키당 결속은 저장 구조에서 사라졌다.
