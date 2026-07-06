# UnivDB 향후 방향 구체화: Phase 6.5 ~ R5 실행 계획

작성일: 2026-07-05
상위 문서: `../UnivDB_refactor_plan.md` (Phase 0~8 원 설계), `../UnivDB_design_decisions_202607.md` (설계 결정 v2: 세대 계약 / 컬럼 접근 밀봉)
근거: 설계 결정 메모 9항 (내부 언어 전환 / Dictionary 밀봉 / Core 커널 유지 / rowID 중심 실행)

---

## 1. 설계 결정 대비 현재 상태

| 결정 | 상태 | 비고 |
|---|---|---|
| 1. 내부 언어를 HashIndex/RowList/RowSet/IndexCursor로 | **~90% 완료** | Where/Filter/Join 신경로 완료. GroupBy만 구 스타일 누적기 잔존 |
| 2. Dictionary 밀봉 (폐기 아님) | **완료** | raw Dictionary 인덱스 경로 제거, public shim은 HashIndex 어댑터로 유지 |
| 3. Core는 저수준 커널 유지 | **준수 중** | Core에 DB 의미론 침범 없음. CoreHashIndex는 DB layer adapter |
| 4. Join/Where/GroupBy/OrderBy는 DB 레이어 | **준수 중** | 모두 UnivDB 레이어에 있음. GroupBy는 아직 구 스타일(로컬 Dictionary 누적기) |
| 5. rowID 중심, 파사드 경계 물질화 | **완료** | 쿼리 경로 중간 물질화 없음. SelectTable/ToArray 경계 확립 |
| 6. 폐기 아닌 엔진화 | **진행 중** | Phase 0~8 완료, 다음 병목은 GroupBy/컬럼 캐시 |
| 7. 작은 검증 단위 커밋 | **준수 중** | vh run / vh check 루프 유지 |
| 8. CoreHashIndex는 후속 선택지 | **완료** | 기본은 DictHashIndex 유지. CoreHashIndex는 build/join/cursor-heavy 명시 옵션 |

---

## 2. 현재 갭 재분류

### G1. 레거시 raw Dictionary 인덱스 경로 — 해소됨
`clsTable`의 인덱스 진실은 `m_dicHashIndices` / `DbHashIndex`로 단일화되었다.

- `SelectRows` / `SelectRowsByIndex`는 `DbRowList.ToCollection` 기반 호환 shim이다.
- `GetUniqueKeys` / `GetUniqueKeysByIndex`는 `DbIndexCursor` 순회 기반 호환 shim이다.
- `UpsertFrom`도 `GetHashIndex` + `TryGetRows` 경로를 사용한다.

현재 action 없음. 이후 금지사항은 raw `Dictionary(key -> Collection)` 인덱스 경로 재도입 금지다.

### G2. 백엔드 교체 불가능한 타입 시그니처 — 해소됨
`DbHashIndex` 인터페이스가 도입되었고, `DictHashIndex` / `CoreHashIndex`가 같은 표면 뒤에 있다.
`clsTable.GetHashIndex(ByIndex)`, `clsQuery`, `modEngine.JoinTable`은 `DbHashIndex`를 소비한다.

### G3. GroupBy 키 의미론 불일치
`clsTable.GroupBy` (clsTable.cls:509~676)는 로컬 Dictionary 누적기 + `CStr & "|"` 복합 키를 쓴다. 노출 누수는 없으나(로컬 변수), `CompareMode` 미설정으로 **대소문자 구분(binary)** — 인덱스 계약(vbTextCompare, 대소문자 무시)과 다르다. `"IT"`와 `"it"`가 Where에서는 같은 키, GroupBy에서는 다른 그룹이 된다. 의도된 동작인지 결정 필요 (아래 D2).

### G4. Core 측 프리미티브 — 부분 해소
- `Core_clsFastHash`와 `Core_clsLongList`는 UnivCore lock 기준으로 공급된다.
- `CoreHashIndex`는 slot→payload(`DbRowList`) 연결을 DB layer 병렬 배열로 해결한다.
- `Core_clsLongSet` 또는 rowset 전용 Core primitive는 아직 보류다. 현재 병목상 우선순위는 GroupBy/컬럼 캐시 쪽이다.

### G5. UnivDB ↔ UnivCore 통합 메커니즘 — 해소됨
`deps.lock` + `vh deps` hydrate 방식으로 전환했다.
`src/Core_*`는 설치 산출물(gitignore)이고, UnivDB는 lock 파일로 UnivCore 버전과 SHA-256을 고정한다.

---

## 3. 실행 계획

아래 Phase 6.3~8은 완료된 이력이다. 다음 실제 진행 순서는 `Phase R0`부터 따른다.

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

결정(2026-07-06 1차 벤치):
- 기본 백엔드는 `DictHashIndex`로 유지한다.
- `CoreHashIndex`는 `clsTable.IndexBackend = dbBackendCore` 명시 옵션으로만 사용한다.
- 근거: 20k행 기준 Core는 index build에서만 개선(low-dup 58→38ms, high-dup 91→82ms)이고, Where/Join 주요 핫패스는 Dict와 동등하거나 더 느리다(Where hit 96→98ms, miss 95→101ms, Join 1:1 29→37ms, no-match left 43→63ms). 30% 개선/2배 핫패스 조건을 충족하지 못했다.

보정(2026-07-06 cursor 최적화 후):
- `CoreHashIndex.Cursor`에서 임시 `Scripting.Dictionary` 재물질화를 제거한 뒤, Core는 join/cursor-heavy 경로에서 유의미한 이득을 보였다.
- 최신 관측: join 1:1 약 19% 개선, no-match-heavy left join 약 18% 개선, low-dup index build 약 36% 개선.
- Where 단독 조회는 여전히 Dict와 동등하거나 소폭 불리하다.
- 따라서 정책은 다음으로 보정한다:
  - 전역 기본은 `DictHashIndex` 유지.
  - `CoreHashIndex`는 폐기하지 않고 build-heavy / join-heavy / cursor-heavy 워크로드의 명시 옵션으로 유지.
  - 자동 선택은 아직 금지. GroupBy 엔진화 후 backend policy v2에서 재평가한다.

커밋: `perf: evaluate hash index backends and set default policy`
후속 커밋: `perf: avoid dictionary materialization for core index cursor`

---

## 3.1 재정렬된 다음 Phase

### Phase R0 — 문서/상태 정렬

목표:
- 완료된 Phase 6.3~8과 실제 코드 상태를 문서에 반영한다.
- CoreHashIndex 정책을 cursor 최적화 이후의 벤치 기준으로 보정한다.
- 다음 작업을 GroupBy / 컬럼 캐시 / backend policy v2 축으로 재배치한다.

범위:
- 구현 변경 없음.
- `tasks/2026-07-05_next_direction.md` 상태/갭/phase 정리.

검증:
- 문서 diff만 확인.

### Phase R1 — GroupBy 의미론 결정 테스트

목표:
- GroupBy key 의미론을 배포 전 API 계약으로 확정한다.

결정 후보:
- A. 현행 유지: GroupBy는 대소문자 구분(binary), Where/index는 대소문자 무시.
- B. 인덱스 계약으로 통일: GroupBy도 `NormalizeIndexKey` / case-insensitive / Empty→"" / Error skip 의미론을 따른다.

권장:
- B. Where / Join / GroupBy의 key 의미론을 같은 DB 계약으로 통일한다.

작업:
- 현재 동작을 드러내는 회귀 테스트를 먼저 추가한다.
- 결정 후 기대값을 바꾸는 별도 커밋으로 처리한다.
- 복합 key는 별도 helper로 분리하기 전까지 string concat 정책의 충돌 가능성을 문서화한다.

### Phase R2 — 단일 컬럼 GroupBy 엔진화

목표:
- 단일 컬럼 GroupBy를 `DbHashIndex.Cursor` 기반으로 재작성한다.
- `key -> RowList`가 이미 group partition이므로, 전체 행 재스캔과 로컬 Dictionary 누적기를 줄인다.

작업:
- 단일 group column + 단순 aggregate 경로부터 분기한다.
- `idx.Cursor` 순회로 그룹 key와 row list를 얻고, row list 안에서 aggregate column만 읽는다.
- Dict/Core 양쪽 벤치 비교를 기록한다.

검증:
- 기존 GroupBy 테스트 green.
- GroupBy short/long key 벤치 비교.
- Core cursor 최적화가 GroupBy에도 이득을 주는지 확인.

### Phase R3 — 복합 GroupBy key 설계

목표:
- 복합 key를 단순 `CStr & "|"` concatenation에서 분리한다.

작업 후보:
- `DbCompositeKey` helper 또는 `modIndexKey` 내 복합 key encode helper 도입.
- delimiter collision, Empty, Error, case-insensitive 비교 정책을 명시한다.
- 단일 컬럼 엔진화 결과를 보고 복합 컬럼까지 HashIndex화할지 결정한다.

### Phase R4 — ColumnReader / hot column cache 재평가

목표:
- 현재 `CopyColumn1D` 기반 컬럼 접근 밀봉 위에 실제 캐시를 둘 가치가 있는지 측정한다.

작업:
- 같은 컬럼을 반복 읽는 Filter / GroupBy / index build 케이스를 분리 벤치한다.
- 세대 계약에 묶인 hot column cache 후보를 설계한다.
- CoreHashIndex보다 큰 지렛대인지 비교한다.

### Phase R5 — Backend policy v2

목표:
- GroupBy 엔진화와 hot column cache 벤치 이후 백엔드 정책을 재평가한다.

현재 기본값:
- `DictHashIndex` 기본 유지.
- `CoreHashIndex`는 명시 옵션 유지.

재평가 기준:
- Where / Join / GroupBy 전체에서 Core가 안정적으로 30% 이상 이기면 기본 전환 검토.
- 특정 경로만 이기면 자동 선택 대신 public/내부 명시 옵션 유지.
- 자동 backend 선택은 cost model 없이 도입하지 않는다.

---

### 과거 Phase 9 후보 — GroupBy 엔진화

Phase 8까지의 결과를 보고 판단. 그룹 키가 단일 컬럼이면 기존 `DbHashIndex`(키→RowList)가 이미 그룹핑 그 자체이므로, `idx.Cursor` 순회 + 집계로 재작성 가능(인덱스 재사용 시 재스캔 제거). 복합 키는 별도 결정 필요. **지금 단계의 스코프가 아님** — 문서화만 해둔다.

이 항목은 `Phase R1~R3`으로 분해되었다.

---

## 4. 미결 결정 (착수 전 확인)

- **D1. Core 통합 방식**: ~~모듈 복사 동기화(권장) vs xlam 참조~~ → **결정됨 (2026-07-06)**:
  릴리스 패키지 + `vh deps` 잠금 배포 채택. 문제의식은 UnivDB와 UnivFilter가 UnivCore를
  공유해야 하는가였고, 답은 "소스 복사가 아니라 버전 잠금 채널로 공유"다.
  - 생산자: UnivCore가 `Core 패키지 내보내기`(core.manifest 기준)로 `releases/UnivCore_latest.txt` 발행
  - 소비자: `vh deps install`로 `src/Core_*` 전개(gitignore) + `deps.lock`(버전+SHA-256) 커밋,
    `vh run`이 lock 기준 자동 hydrate. UnivDB는 UnivCore@20260706_1746(FastHash 포함) 채택 완료
  - UnivFilter는 현재 구 릴리스(20260705_1005)와 일치하는 수동 사본 - 같은 채널로 후속 전환 (보류 중)
  - xlam 참조 방식은 원안대로 비채택
- **D2. GroupBy 키 의미론**: 대소문자 구분(현행) 유지 vs 인덱스 계약(무시)과 통일. 통일 시 동작 변경이므로 별도 커밋 + 테스트 명시 필요. Phase 9 전까지는 현행 유지. (설계 결정 v2 4장의 "비결정 D: NULL/타입 의미론"의 하위 항목 — 라이브러리 1.0 공개 전 확정 필수)

---

## 5. 금지 사항 (원 계획 8장 재확인)

- Core에 DB 의미론 주입 금지 (`Core_JoinTables`, `Core_GroupBy` 류)
- 쿼리 실행 경로에 `someDict.Keys` / `New Collection` / `clsSet` 재도입 금지
- public 호환 shim(`SelectRows` 등) 시그니처 변경 금지
- 백엔드 자동 선택 조기 도입 금지
