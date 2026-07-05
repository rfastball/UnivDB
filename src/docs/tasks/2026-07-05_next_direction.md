# UnivDB 향후 방향 구체화: Phase 6.5 ~ 8 실행 계획

작성일: 2026-07-05
상위 문서: `../UnivDB_refactor_plan.md` (Phase 0~8 원 설계), `../UnivDB_design_decisions_202607.md` (설계 결정 v2: 세대 계약 / 컬럼 접근 밀봉)
근거: 설계 결정 메모 9항 (내부 언어 전환 / Dictionary 밀봉 / Core 커널 유지 / rowID 중심 실행)

---

## 1. 설계 결정 대비 현재 상태

| 결정 | 상태 | 비고 |
|---|---|---|
| 1. 내부 언어를 HashIndex/RowList/RowSet/IndexCursor로 | **~80% 완료** | Where/Filter/Join 신경로 완료. 레거시 raw Dictionary 인덱스 경로 잔존 |
| 2. Dictionary 밀봉 (폐기 아님) | **부분 완료** | DictHashIndex 밀봉 완료. clsTable 이중 인덱스 캐시가 남은 누수원 |
| 3. Core는 저수준 커널 유지 | **준수 중** | Core에 DB 의미론 침범 없음. 단 LongList/LongSet 공개 프리미티브 미비 |
| 4. Join/Where/GroupBy/OrderBy는 DB 레이어 | **준수 중** | 모두 UnivDB 레이어에 있음. GroupBy는 아직 구 스타일(로컬 Dictionary 누적기) |
| 5. rowID 중심, 파사드 경계 물질화 | **완료** | 쿼리 경로 중간 물질화 없음. SelectTable/ToArray 경계 확립 |
| 6. 폐기 아닌 엔진화 | **진행 중** | Phase 0~6 커밋 완료 (78bfb54 ~ cad338c) |
| 7. 작은 검증 단위 커밋 | **준수 중** | vh run / vh check 루프 유지 |
| 8. CoreHashIndex는 후속 선택지 | **미착수** | 병목 측정(벤치)조차 아직 없음 — 선행 조건부터 |

---

## 2. 탐색으로 확인된 잔여 갭

### G1. 레거시 raw Dictionary 인덱스 경로 (최우선)
`clsTable`이 인덱스 캐시를 **이중으로** 유지한다.

- 구경로: `GetIndex` / `GetIndexByIndex` (clsTable.cls:269, 298) → raw `Dictionary(key -> Collection)` + `m_dicIndices` 캐시
- 신경로: `GetHashIndex` / `GetHashIndexByIndex` (clsTable.cls:342, 365) → `DictHashIndex` + `m_dicHashIndices` 캐시

구경로 소비자 (전환 대상):
- `SelectRows` (clsTable.cls:391) / `SelectRowsByIndex` (:409) — 호환 shim, Collection 반환 유지 필요
- `GetUniqueKeys` (:425) / `GetUniqueKeysByIndex` (:437) — 호환 shim
- `UpsertFrom` (:758) — **내부 실행 경로인데 구경로 사용 중** (shim 아님)

같은 컬럼 인덱스가 두 번 빌드·캐시될 수 있고, 무효화 정책도 두 캐시를 모두 관리해야 하는 부담이 있다.

### G2. 백엔드 교체 불가능한 타입 시그니처
`GetHashIndex(...) As DictHashIndex` — 반환 타입이 **구체 클래스**다. Phase 7에서 CoreHashIndex를 같은 표면 아래 넣으려면 `DbHashIndex` 인터페이스(VBA `Implements`) 도입이 선행돼야 한다. 소비자는 clsQuery.cls:67, modEngine.bas:67-68 두 곳뿐이라 전환 비용이 작은 지금이 적기다.

### G3. GroupBy 키 의미론 불일치
`clsTable.GroupBy` (clsTable.cls:509~676)는 로컬 Dictionary 누적기 + `CStr & "|"` 복합 키를 쓴다. 노출 누수는 없으나(로컬 변수), `CompareMode` 미설정으로 **대소문자 구분(binary)** — 인덱스 계약(vbTextCompare, 대소문자 무시)과 다르다. `"IT"`와 `"it"`가 Where에서는 같은 키, GroupBy에서는 다른 그룹이 된다. 의도된 동작인지 결정 필요 (아래 D2).

### G4. Core 측 프리미티브 미비
- `Core_clsLongList` / `Core_clsLongSet` 공개 프리미티브 없음 (결정 3항에서 Core 책임으로 명시)
- `Core_clsFastHash`는 키 전용 — slot→payload(DbRowList) 연결은 DB 레이어가 slot index 기반 병렬 배열로 해결 가능 (Core 변경 불필요, 단 build-only 사용 전제. FastHash의 swap-last delete는 slot index를 불안정하게 하므로 인덱스 용도에서는 delete 미사용)

### G5. UnivDB ↔ UnivCore 통합 메커니즘 부재
현재 UnivDB에서 `Core_*` 참조 0건. VBA는 패키지 시스템이 없으므로:
- **권장: 모듈 복사 동기화** — `Core_*.bas/cls`를 UnivDB src로 복사 (LogShim과 같은 방식). 하네스가 프로젝트별 src를 import하는 구조와 일치
- 대안(비권장): xlam 애드인 참조 — 클래스 인스턴스화에 PublicNotCreatable + 팩토리 필요, 배포·버전 결합 부담

---

## 3. 실행 계획

### Phase 6.3 — 세대(Generation) 계약 도입 (설계 결정 v2-A, 선행 축)

rowID를 "물리 행 번호 + 세대 계약"으로 명시화한다. 모든 파생 구조의 무효화를 하나의 규칙으로 통일하는 기반이다.

1. `clsTable`에 `Generation As Long` 도입. 모든 구조 변경 경로(Insert/Delete/OrderBy/Upsert/컬럼 추가·삭제)에서 증가
2. `DictHashIndex` / `DbRowSet` / `DbRowList`가 생성 시점의 (테이블, 세대)를 기억하고, 세대 불일치 상태 사용 시 `Err.Raise` (fail-fast)
3. 기존 인덱스 무효화 helper(Phase 1 산출물)를 세대 규칙의 특수 사례로 흡수 — 세대 증가가 곧 전체 파생물 실효
4. 테스트: 구조 변경 후 이전 RowSet/인덱스 사용이 에러가 되는지, 정상 경로는 회귀 없는지

커밋: `feat: bind derived structures to table generation`
검증: 기존 전 테스트 green + stale 사용 fail-fast 테스트 신규 통과.

### Phase 6.5 — 레거시 인덱스 경로 단일화 (G1)

인덱스의 유일한 진실은 `DictHashIndex` 하나가 되게 한다.

1. **6.5a** `UpsertFrom`을 `GetHashIndex` + `TryGetRows` 기반으로 전환 (내부 실행 경로 정리)
   - 커밋: `refactor: route upsert through hash index`
2. **6.5b** `SelectRows` / `SelectRowsByIndex` / `GetUniqueKeys(ByIndex)` shim을 DictHashIndex 어댑터로 재구현 (`DbRowList.ToCollection`, `Cursor` 순회로 key 배열 생성). 이후 `GetIndex` / `GetIndexByIndex` / `m_dicIndices` 제거
   - public 시그니처(Collection/Variant 반환)는 그대로 — 결정 2.4 호환 계약 유지
   - 커밋: `refactor: retire raw dictionary index path`
3. **6.5c** `clsQuery.Delete`의 로컬 `dicToDelete` → `DbRowSet.ContainsRow` (저강도, 겸사겸사)
   - 커밋: 6.5b에 포함 가능

검증: 기존 전 테스트(특히 TestBaseline, TestIndexInvalidation, TestHashIndex의 SelectRows 패리티) green. 무효화 경로가 단일 캐시만 관리하는지 확인.

### Phase 6.7 — DbHashIndex 인터페이스 도입 (G2)

1. `DbHashIndex.cls` 인터페이스 클래스 신설: `AddRow / TryGetRows / CountKeys / Clear / Cursor`
2. `DictHashIndex`에 `Implements DbHashIndex` 추가 (기존 public 멤버는 위임)
3. `clsTable.GetHashIndex(ByIndex)` 반환 타입, `clsQuery` / `modEngine` 소비 지점을 `DbHashIndex`로 전환
4. key 정규화(Empty→"", Error skip, CStr, vbTextCompare)는 백엔드 독립 공통 계약이므로 `DictHashIndex` private에서 공용 helper(`modIndexKey.bas` 등)로 승격 — CoreHashIndex가 동일 계약을 공유할 기반

커밋: `refactor: introduce hash index interface surface`
검증: 전 테스트 green + 성능 회귀 없음 (VBA Implements 디스패치 오버헤드 벤치 — Phase 7 준비 벤치로 겸용).

### Phase 6.8 — 컬럼 접근 밀봉 (설계 결정 v2-C)

정본 스토리지(row-major 2D Variant)는 유지하되, 실행계가 `m_vData`를 직접 만지지 못하게 컬럼 접근 표면 뒤로 밀봉한다.

1. 컬럼 접근 표면 도입 (`ColumnReader` 또는 `clsTable`의 컬럼 순회 메서드) — Where/Filter/GroupBy/인덱스 빌드의 셀 접근을 이 표면으로 수렴
2. 표면은 세대 계약(Phase 6.3)의 파생물 규칙을 따름 — 세대 불일치 시 fail-fast
3. 이 표면 뒤에서 **핫 컬럼 캐시**(병목 컬럼만 1D 배열로 지연 물질화)를 도입할 수 있는 구조 확보 — 캐시 자체 도입은 Phase 7 벤치 결과를 보고 결정

커밋: `refactor: seal column access behind reader surface`
검증: 전 테스트 green + 밀봉 후 스캔 성능 회귀 없음 (벤치 비교).

### Phase 7 준비 — 벤치마크 스위트 (결정 8의 "병목 측정" 선행 조건)

`TestBenchDb.bas` 신설, 원 계획 7.4 목록 기준:
- index build N행 / Where equality hit·miss·고중복·저중복 / 다중 Where 교집합 / Filter 순회 / Join 1:1·1:N·N:N / no-match-heavy left join / GroupBy 짧은·긴 키
- **지렛대 비교: 핫 컬럼 캐시(Phase 6.8 표면 뒤 1D 캐시) vs CoreHashIndex** — 어느 쪽이 더 큰 성능 지렛대인지 측정해 Phase 7/캐시 도입의 우선순위 결정
- warm-up 후 다회 측정, median 기록, `LogJ`로 jsonl에 남겨 baseline 보존

커밋: `test: add db-layer benchmark suite and record baseline`

### Phase 7 — CoreHashIndex 실험 백엔드 (G4, G5)

선행: UnivCore에 `Core_clsLongList` 추가(Append/Count/ItemAt/ToArray 최소 표면), Core 테스트 통과 후 UnivDB src로 `Core_clsFastHash` + 의존 모듈(`Core_modMemCore`, `Core_modCompare`, `Core_modNorm`, `Core_Enums`) 복사 동기화.

구현 (UnivDB 레이어):
```text
CoreHashIndex (Implements DbHashIndex)
  Private mEngine As Core_clsFastHash     ' 키 저장·slot 결정
  Private mRows() As DbRowList            ' slot index -> DbRowList (병렬 배열)
  key 정규화: Phase 6.7의 공통 helper 사용 (백엔드 동치 계약)
  delete 미지원 (인덱스는 build-only, Clear/Rebuild만)
```

검증: **TestHashIndex를 백엔드 파라메트릭으로 개조** — 동일 케이스를 DictHashIndex/CoreHashIndex 양쪽에 실행해 동치 확인 (Empty, 대소문자, 숫자 키, Error skip, 중복 키). 이후 Phase 7 준비 벤치를 CoreHashIndex로 재실행해 비교.

커밋: `feat: add experimental core-backed hash index`

### Phase 8 — 백엔드 선택 정책 (원 계획 유지)

- 향상 < 30% → DictHashIndex 기본 유지
- 특정 핫패스 2배 이상 → 해당 경로/키 프로파일에 한해 CoreHashIndex 검토
- 자동 선택은 도입하지 않음 — 명시적 옵션 + 벤치 기반 수동 판단

커밋: `perf: evaluate hash index backends and set default policy`

### Phase 9 (후속 선택지) — GroupBy 엔진화

Phase 8까지의 결과를 보고 판단. 그룹 키가 단일 컬럼이면 기존 `DbHashIndex`(키→RowList)가 이미 그룹핑 그 자체이므로, `idx.Cursor` 순회 + 집계로 재작성 가능(인덱스 재사용 시 재스캔 제거). 복합 키는 별도 결정 필요. **지금 단계의 스코프가 아님** — 문서화만 해둔다.

---

## 4. 미결 결정 (착수 전 확인)

- **D1. Core 통합 방식**: 모듈 복사 동기화(권장) vs xlam 참조. 복사 시 동기 스크립트/절차를 어디에 둘지 (UnivCore releases → UnivDB src).
- **D2. GroupBy 키 의미론**: 대소문자 구분(현행) 유지 vs 인덱스 계약(무시)과 통일. 통일 시 동작 변경이므로 별도 커밋 + 테스트 명시 필요. Phase 9 전까지는 현행 유지. (설계 결정 v2 4장의 "비결정 D: NULL/타입 의미론"의 하위 항목 — 라이브러리 1.0 공개 전 확정 필수)

---

## 5. 금지 사항 (원 계획 8장 재확인)

- Core에 DB 의미론 주입 금지 (`Core_JoinTables`, `Core_GroupBy` 류)
- 쿼리 실행 경로에 `someDict.Keys` / `New Collection` / `clsSet` 재도입 금지
- public 호환 shim(`SelectRows` 등) 시그니처 변경 금지
- 백엔드 자동 선택 조기 도입 금지
