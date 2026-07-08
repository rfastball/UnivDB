# UnivDB 대규모 리팩토링 지침: 내부 실행 표현 전환

## 0. 문서 목적

이 문서는 UnivDB 코드베이스를 대규모로 리팩토링할 코드 에이전트에게 전달하기 위한 설계 지침이다. 목표는 세부 구현을 미리 고정하는 것이 아니라, 현재 프로토타입의 구조적 한계를 정리하고, 앞으로의 내부 실행 모델과 구현 순서를 명확히 정하는 것이다.

핵심 방향은 다음과 같다.

- 현재의 `Dictionary` / `Collection` / `clsSet` 중심 고수준 컨테이너 전달 구조를 쿼리 실행 내부에서 점진적으로 제거한다.
- DB 내부 표준 표현을 `HashIndex` / `RowList` / `RowSet` 중심으로 전환한다.
- 첫 구현은 안정적인 `Dictionary` 백엔드를 사용하되, Dictionary는 반드시 `DictHashIndex` 내부에 밀봉한다.
- 이후 동일 인터페이스 아래 `CoreHashIndex` 백엔드를 실험 가능하게 만든다.
- 구현은 반드시 작게 나누고, 각 단계마다 테스트와 벤치마크를 통과한 뒤 커밋한다.

이 문서는 구현 세부 코드보다 설계 결정, 책임 경계, 마이그레이션 순서, 검증 기준에 집중한다.

---

## 1. 현재 상태 요약

현재 UnivDB 프로토타입은 대략 다음 구조를 사용한다.

```text
Table storage:
  2D Variant array

Headers / catalog:
  Dictionary

Column index:
  Dictionary(key -> Collection(rowIDs))

Candidate rows:
  clsSet

Set operations:
  clsSet(Dictionary-backed)

Join key comparison:
  clsSet.Compare / Intersection

Final output:
  2D Variant array / Table facade / Range output
```

이 구조는 프로토타입으로는 적절했다. 이유는 다음과 같다.

- VBA에서 빠르게 기능을 만들 수 있다.
- `Dictionary`로 key lookup을 쉽게 구현할 수 있다.
- `Collection`으로 rowID 목록을 쉽게 담을 수 있다.
- `clsSet`으로 교집합/차집합을 읽기 쉬운 형태로 표현할 수 있다.
- Excel Range와 2D Variant 배열이 자연스럽게 연결된다.

그러나 쿼리 실행 내부에서 다음 문제가 드러난다.

```text
rowID
  -> Collection
  -> clsSet
  -> KeysArray
  -> 다시 rowID loop
```

조인 경로에서는 다음처럼 key가 불필요하게 왕복한다.

```text
index keys
  -> clsSet
  -> common keys
  -> 다시 index lookup
  -> Collection rows
```

이 문제는 전체 구조의 치명적 결함이라기보다, 프로토타입 특유의 읽기 쉬운 고수준 컨테이너 전달 방식이 성능과 확장성의 한계에 닿은 것으로 본다.

---

## 2. 핵심 설계 결정

### 2.1 지금 결정할 것은 “Dictionary냐 Core냐”가 아니다

가장 중요한 결정은 백엔드 구현체가 아니라 DB 내부 실행 언어다.

기존 내부 언어:

```text
Dictionary
Collection
clsSet
Variant key arrays
```

새 내부 언어:

```text
HashIndex
IndexCursor / IndexEntry
RowList
RowSet
Long rowID
```

따라서 1차 목표는 다음이다.

> QueryEngine, Table, Join, Where 등의 DB 실행 경로에서 `Dictionary`, `Collection`, `clsSet`이 직접 보이지 않게 한다.

그 다음, `HashIndex`의 내부 구현체로 다음 두 백엔드를 둘 수 있게 한다.

```text
DictHashIndex:
  Dictionary 기반 안정 구현

CoreHashIndex:
  Core/FastHash 기반 실험 또는 핫패스 구현
```

### 2.2 Dictionary는 금지 대상이 아니라 밀봉 대상이다

`Dictionary`는 catalog, headers, metadata, 작은 lookup에는 계속 유용하다. 문제는 `Dictionary` API가 쿼리 실행 내부에 직접 퍼지는 것이다.

쿼리 엔진에서 금지할 것:

```vba
someDict.Keys
someDict(key)
someDict.Exists(key)
Collection(rowIDs)
clsSet.Compare for row candidates
```

대신 쿼리 엔진은 다음 추상화만 사용한다.

```vba
idx.TryGetRows key, rows
Set cur = idx.Cursor
candidate.IntersectWith rows
EmitJoin leftRows, rightRows
```

### 2.3 Core는 DB를 모르는 저수준 관계 연산 커널로 유지한다

Core의 책임 범위:

```text
- 순수 Set 연산
- LongList / LongSet / RowID 집합 primitive
- key -> slot 해시 primitive
- hash / equality / cached hash
- slot/view/materialize-on-write 같은 저수준 표현
```

Core가 책임지면 안 되는 것:

```text
- Table / Column / Row 의미론
- SQL NULL 의미론
- index invalidation 정책
- Join / Where / GroupBy / OrderBy 실행 계획
- 결과 2D 배열 생성
- Range 출력
```

DB 레이어는 Core를 사용할 수 있지만, Core가 DB를 알아서는 안 된다.

### 2.4 public API는 호환 shim으로 보호한다

내부 실행 경로와 외부 호환 표면을 분리한다.

기존 코드와 테스트가 기대하는 public 표면:

```text
clsTable.SelectRows(...) As Collection
clsTable.SelectRowsByIndex(...) As Collection
clsTable.GetUniqueKeys(...) As Variant
clsTable.GetUniqueKeysByIndex(...) As Variant
```

이 표면은 리팩토링 중 즉시 제거하지 않는다. 필요한 경우 `DbRowList` / `DbHashIndex`에서 기존 `Collection` 또는 key 배열을 만들어 반환하는 compatibility shim으로 유지한다.

단, 신규 DB 실행 경로의 원칙은 다음이다.

```text
Query / Filter / Join / Update / Delete internals:
  DbHashIndex / DbRowList / DbRowSet만 사용

Legacy public helpers / tests / external callers:
  필요 시 Collection / Variant key array 호환 반환 허용
```

### 2.5 key 정규화와 비교 의미론은 백엔드 독립 계약이다

현재 인덱스 의미론은 다음에 의존한다.

```text
- Dictionary CompareMode = vbTextCompare
- key comparison is case-insensitive for strings
- Empty key is normalized to ""
- Error values are not indexed
- non-error key values are compared through current CStr-based behavior
```

`DictHashIndex`와 `CoreHashIndex`는 이 의미론을 동일하게 지켜야 한다. 따라서 key normalization / equality / hash 정책은 각 백엔드 내부에 흩어두지 않고, DB index layer의 공통 계약으로 둔다.

초기 구현은 별도 객체를 만들지 않고 private helper여도 된다. 그러나 public 동작과 백엔드 동치 테스트의 기준은 반드시 위 규칙이다.

---

## 3. 목표 아키텍처

### 3.1 레이어 구조

```text
Facade layer
  clsDatabase
  clsTable public API
  clsQuery public API
  Report / ToArray / ToSheet

Execution layer
  QueryEngine / clsQuery internals
  Where / Filter / Join / GroupBy / OrderBy
  RowSet operations

DB storage/index layer
  DbHashIndex
  DictHashIndex
  optional CoreHashIndex
  DbRowList
  DbRowSet
  DbIndexCursor

Core layer
  Core_clsSet
  Core_clsLongList
  Core_clsLongSet or equivalent
  Core_clsHashSlots or Core_clsFastHash-derived primitive
  Core_modCompare / Core_modNorm / Core_modMemCore
```

### 3.2 내부 데이터 흐름 원칙

실행 내부에서는 실제 행이나 2D 배열을 조기에 만들지 않는다.

권장 흐름:

```text
WHERE condition
  -> HashIndex probe
  -> RowList / RowSet
  -> RowSet combine
  -> final materialization only at facade boundary
```

조인 흐름:

```text
smallIndex cursor
  -> entry.Key / entry.Rows
  -> largeIndex.TryGetRows(entry.Key)
  -> leftRows x rightRows emit
  -> final result array only when necessary
```

물질화 경계:

```text
- ToArray
- ToTable
- ToRange / WriteToSheet
- external row/record facade
- public API boundary where actual values are requested
```

---

## 4. 새 핵심 추상화

아래 이름은 제안이다. 기존 naming convention에 맞게 조정해도 된다. 단, 책임 경계는 유지한다.

### 4.1 DbRowList

목적: `Collection(rowIDs)` 대체.

개념:

```text
DbRowList
  Append(rowId As Long)
  Count As Long
  RowAt(i As Long) As Long
  ToArray() As Variant or Long()
  Clear
```

설계 의도:

- rowID는 `Variant`가 아니라 `Long`으로 다룬다.
- rowID 목록은 객체 컬렉션이 아니라 배열 기반 append list로 유지한다.
- Dictionary 백엔드와 Core 백엔드 모두 같은 `DbRowList`를 payload로 사용할 수 있다.

### 4.2 DbRowSet

목적: 후보 행 집합을 `clsSet` 대신 표현.

개념:

```text
DbRowSet
  FromAllRows(tableRowCount)
  FromRowList(rows)
  IntersectWith(rows or rowSet)
  UnionWith(rows or rowSet)
  DifferenceWith(rows or rowSet)
  ContainsRow(rowId As Long)
  Cursor / ToArray
```

초기 구현은 단순해도 된다.

가능한 내부 전략:

```text
- sorted Long array
- unsorted LongList + mark bitmap
- Boolean bitmap
- Core LongSet wrapper
```

중요한 점:

- QueryEngine은 더 이상 `clsSet`을 row 후보 표현으로 사용하지 않는다.
- `DbRowSet`은 특정 Table의 rowID 의미를 가진 DB 레이어 객체다.
- 정렬 없는 `SelectTable`, `Top`, `Update`, `Delete`의 호환을 위해 cursor / ToArray 결과는 원본 rowID 오름차순을 기본 순서로 한다.
- 내부 전략이 bitmap 또는 hash set이더라도 외부로 순회될 때는 안정적인 rowID 오름차순을 보장한다.

### 4.3 DbHashIndex

목적: 컬럼 인덱스의 공통 표면.

개념:

```text
DbHashIndex
  AddRow(key, rowId)
  TryGetRows(key, outRows) As Boolean
  Cursor() As DbIndexCursor
  CountKeys As Long
  NormalizeKey(key) or equivalent shared key policy
  Version / SourceColumn metadata if needed
  Clear / Rebuild
```

중요한 규칙:

- QueryEngine은 Dictionary를 보지 않는다.
- QueryEngine은 Collection을 보지 않는다.
- QueryEngine은 `DbHashIndex`와 `DbRowList`만 본다.
- 모든 백엔드는 동일한 key 정규화, Error skip, Empty 처리, case-insensitive 비교 의미론을 제공한다.

### 4.4 DbIndexCursor / DbIndexEntry

목적: `Dictionary.Keys` 노출 방지.

개념:

```text
DbIndexCursor
  MoveNext() As Boolean
  Key As Variant
  Rows As DbRowList
```

또는 entry 객체를 반환해도 된다.

```text
DbIndexEntry
  Key
  Rows
  optional SlotId
  optional Hash
```

초기 `DictHashIndex`는 내부적으로 key cache나 `Dictionary.Keys`를 쓸 수 있다. 단, 그것은 Cursor 내부에서 끝나야 한다.

---

## 5. 백엔드 전략

### 5.1 1차 백엔드: DictHashIndex

초기 목표는 안정적인 Dictionary 기반 구현이다.

내부 구조 예:

```text
DictHashIndex
  Private mDict As Scripting.Dictionary
    key -> DbRowList
```

선택적으로 key cache를 둘 수 있다.

```text
mKeys()
mDirtyKeys
mVersion
```

그러나 key cache 최적화는 1차 구현의 필수 목표가 아니다. 먼저 인터페이스 밀봉과 동작 일치가 우선이다.

### 5.2 2차 백엔드: CoreHashIndex

CoreHashIndex는 즉시 전면 도입하지 않는다. `DbHashIndex` 표면이 안정된 뒤 실험한다.

목표 구조:

```text
CoreHashIndex
  Core hash slots / FastHash-derived storage
  slot -> key
  slot -> hash
  slot -> DbRowList
```

핵심 이점:

- key 대신 slot/hash/payload 중심 실행 가능
- Cursor가 실제 slot 순회에 가까워짐
- Join에서 smallIndex slot을 돌며 largeIndex probe 가능
- repeated lookup / key rematerialization 감소 가능

주의:

- Core는 DB 의미론을 알면 안 된다.
- `key -> rowList` 전체 정책은 DB layer 책임이다.
- CoreHashIndex는 반드시 DictHashIndex와 같은 테스트를 통과해야 한다.

---

## 6. 구현 순서와 커밋 계획

각 단계는 작게 구현하고, 테스트 통과 후 커밋한다. 한 커밋에 여러 계층을 동시에 바꾸지 않는다.

### Phase 0: 기준선 고정

목표:

- 현재 동작을 보호할 테스트와 벤치마크 기준선을 확보한다.

작업:

- 기존 Query / Where / Filter / Join / GroupBy / Update / Delete 동작 테스트 정리.
- 현재 대표 벤치마크 수집.
- 특히 Join, Where chain, Filter, SelectTable, Update 후 재조회 테스트 추가.
- 하네스 루프 기준을 고정한다. 일상 검증은 `src\` 기준 `vh run`, 실패 분석은 `vh log -FailedOnly`, 커밋 전 검사는 `vh check`로 수행한다.
  (2026-07-08 이후 검증 도구는 `vbaflow`로 전환됨: `vbaflow push` + `vbaflow test` / `vbaflow test --json` detail / `vbaflow lint`. 아래 `vh ...`는 당시 이력 표기. 현행은 CLAUDE.md 참조)

검증:

- 기존 기능 테스트 모두 통과.
- 현재 성능 로그 저장.
- 하네스 대상 워크북이 `dev_UnivDB.xlsb`로 유지되는지 확인.

커밋:

```text
test: capture UnivDB behavioral and performance baseline
```

---

### Phase 1: 인덱스 무효화 규칙 정리

목표:

- 백엔드 교체 전, 인덱스 stale 위험을 먼저 줄인다.

작업 방향:

- 인덱스 무효화 helper를 먼저 만든다.

```text
InvalidateIndexByColumnName(colName)
InvalidateIndexByIndex(colIndex)
InvalidateAllIndices()
```

- 컬럼 값 변경 시 해당 컬럼 인덱스 무효화 helper만 사용한다.
- 행 추가/삭제/업데이트 시 관련 인덱스 무효화 또는 재빌드 정책을 helper 기준으로 명확화한다.
- 문자열 컬럼명 key와 숫자 컬럼 index key가 동시에 캐시되는 경우 둘 다 무효화.
- UpdateBatch / Upsert / Value Let / AddColumn / RemoveColumn / Compute / CopyColumn / OrderBy 등 인덱스 영향 경로 점검.
- update-only `UpsertFrom`도 변경된 컬럼 인덱스를 반드시 무효화한다.
- `Value Let`, `CopyColumn`, `Compute`, `AddColumn`, `RemoveColumn`은 현재 캐시가 남을 수 있으므로 Phase 1에서 먼저 막는다.

검증:

- 인덱스 생성 후 값 변경.
- 변경 전 key로 조회되지 않고, 변경 후 key로 조회되는지 확인.
- update-only upsert에서도 index stale이 없는지 확인.
- 컬럼명 기반 캐시와 숫자 index 기반 캐시가 동시에 제거되는지 확인.

커밋:

```text
fix: centralize and harden table index invalidation
```

---

### Phase 2: DbRowList 도입

목표:

- `Collection(rowIDs)`를 DB 전용 Long 기반 row list로 대체할 준비를 한다.

작업 방향:

- `DbRowList` 또는 `Core_clsLongList` 래퍼 도입.
- Append, Count, RowAt, ToArray 정도의 최소 표면만 구현.
- 기존 Collection 반환 경로와 병행 가능하게 adapter를 둔다.

검증:

- RowList append / iterate / count 테스트.
- 기존 SelectRows 결과와 RowList 결과 동일성 검증.

커밋:

```text
feat: add Long-based row list abstraction
```

---

### Phase 3: DictHashIndex 도입 및 Dictionary 밀봉

목표:

- 컬럼 인덱스를 `Dictionary(key -> Collection)`에서 `DictHashIndex(key -> DbRowList)`로 전환한다.
- QueryEngine 외부에 Dictionary와 Collection이 보이지 않게 한다.

작업 방향:

- `DbHashIndex` 표면 도입.
- `DictHashIndex` 구현.
- `clsTable.GetIndexByIndex` 또는 동등 경로가 raw Dictionary가 아니라 DbHashIndex를 반환하도록 변경.
- 기존 public API는 임시 adapter로 유지한다. `SelectRows` / `SelectRowsByIndex`는 외부 호환을 위해 `Collection` 반환을 유지할 수 있으나, QueryEngine 신규 경로는 `DbHashIndex.TryGetRows`를 사용한다.
- `GetUniqueKeys` / `GetUniqueKeysByIndex`도 외부 호환 shim으로 남기되, Join 신규 경로는 key 배열 materialization에 의존하지 않는다.
- key 정규화와 비교 의미론을 `DictHashIndex` 내부 또는 공통 helper에 모아 `CoreHashIndex`와 같은 계약을 공유하게 한다.

검증:

- SelectRows / Where equality / repeated lookup 결과 동일.
- 인덱스 key count 동일.
- Dictionary / Collection이 QueryEngine에 직접 노출되지 않는지 코드 검색.
- `Empty`, 문자열 대소문자, 숫자 key, Error skip 동작이 기존과 동일.

커밋:

```text
refactor: introduce sealed dictionary-backed hash index
```

---

### Phase 4: DbRowSet 도입 및 후보 행 표현 교체

목표:

- 후보 row 집합에서 `clsSet` 의존 제거.

작업 방향:

- `DbRowSet` 도입.
- Query candidate를 `DbRowSet`으로 전환.
- Where는 `DbHashIndex.TryGetRows` 결과를 candidate와 결합.
- Filter는 candidate RowSet을 순회하며 survivor RowSet 생성.
- SelectTable은 RowSet을 최종 물질화 경계로 사용.
- RowSet cursor / ToArray는 rowID 오름차순을 보장한다.

검증:

- Where 단일 조건.
- Where 다중 조건 AND.
- Filter 조합.
- Empty candidate.
- Full table candidate.
- Delete / Update candidate 동작.
- 정렬 없는 SelectTable 결과 순서와 Top 결과가 기존과 동일.

커밋:

```text
refactor: replace candidate clsSet with row-id rowset
```

---

### Phase 5: Join을 direct probe 방식으로 전환

목표:

- Join에서 key set materialization과 재조회 경로 제거.

기존 흐름:

```text
left keys -> clsSet
right keys -> clsSet
common keys -> index lookup again
```

목표 흐름:

```text
leftIndex.Cursor or rightIndex.Cursor
  -> entry.Key / entry.Rows
  -> oppositeIndex.TryGetRows(entry.Key)
  -> emit matched, semi, anti, or outer rows according to join type
```

작업 방향:

- Join sizing pass와 fill pass 모두 DbHashIndex / DbIndexCursor / DbRowList 사용.
- 가능한 경우 같은 key lookup 결과를 sizing/fill에서 재사용할지 검토하되, 1차 구현에서는 명료성을 우선한다.
- outer join 규칙은 DB 레이어에서 유지한다.
- `jtInner`, `jtLeft`, `jtRight`, `jtFull`, `jtLeftAnti`, `jtRightAnti`, `jtLeftSemi`, `jtRightSemi` 전체를 동일 커밋에서 보존한다.
- left-only / right-only 처리는 key set materialization 대신 cursor pass와 matched-key marking으로 구현한다.
- matched-key marking은 DB layer 전용 구조로 두며 `clsSet`을 다시 도입하지 않는다.
- 결과 행 순서는 기존 `JoinTable`과 호환되도록 common keys, left-only, right-only 처리 순서를 명시적으로 유지하거나 테스트로 고정한다.

검증:

- inner join.
- left join.
- right join.
- full join.
- left anti / right anti.
- left semi / right semi.
- 중복 key join.
- no-match join.
- large left / small right, small left / large right.
- 기존 JoinTable 결과와 행 수 및 값 동일.

커밋:

```text
refactor: execute joins through hash index cursors
```

---

### Phase 6: clsSet 역할 축소

목표:

- `clsSet`을 DB 후보 row 처리와 join key 비교에서 제거하고, 순수 값 집합 유틸 또는 외부 호환 계층으로만 남긴다.

작업 방향:

- DB 실행 경로에서 `clsSet.Compare`, `clsSet.KeysArray` 사용 제거.
- 값 집합 유틸이 필요한 경우 명시적으로 Core/Set utility 영역에서만 사용.
- RowID 처리에는 반드시 DbRowSet / DbRowList 사용.

검증:

- 코드 검색으로 DB 실행 경로의 clsSet 의존 확인.
- 기능 회귀 테스트.

커밋:

```text
refactor: confine clsSet to value-set utility usage
```

---

### Phase 7: CoreHashIndex 실험 백엔드 추가

목표:

- DictHashIndex와 같은 표면 아래 CoreHashIndex를 추가해 실제 쿼리 경로에서 비교 가능하게 한다.

작업 방향:

- Core에 DB 의미론을 넣지 않는다.
- CoreHashIndex는 DB layer에 위치한다.
- 내부에서 Core hash primitive를 사용해 key slot을 관리하고, slot별 DbRowList를 보관한다.
- backend selection은 명시적 옵션 또는 내부 실험 플래그로 둔다.

검증:

- DictHashIndex와 동일한 테스트 전체 통과.
- Where / Join / GroupBy relevant benchmark 비교.
- 짧은 문자열, 긴 문자열, 숫자 key, 중복률 높은 key, 중복률 낮은 key 케이스 벤치.

커밋:

```text
feat: add experimental core-backed hash index
```

---

### Phase 8: 백엔드 선택 정책 정리

목표:

- CoreHashIndex를 기본으로 할지, DictHashIndex를 기본으로 할지 실제 벤치 후 결정한다.

판단 기준:

```text
DictHashIndex 대비 CoreHashIndex 향상 < 30%:
  DictHashIndex 기본 유지.

특정 핫패스에서 2배 이상 반복 향상:
  해당 경로 또는 key profile에서 CoreHashIndex 사용 검토.

긴 문자열에서 CoreHashIndex가 불리:
  긴 문자열 key는 DictHashIndex 유지 또는 auto backend 고려.
```

주의:

- backend 자동 선택은 너무 일찍 넣지 않는다.
- 먼저 수동 선택과 벤치 기반 판단을 우선한다.

커밋:

```text
perf: evaluate hash index backends and set default policy
```

---

## 7. 테스트 전략

### 7.1 기능 테스트

필수 케이스:

```text
- 빈 테이블
- 단일 행 테이블
- 중복 key 인덱스
- 없는 key 조회
- Null / Empty / Error 처리 정책
- 문자열 대소문자 무시 key 조회
- 숫자 key와 문자열 key의 기존 CStr 기반 호환
- Where 단일 조건
- Where 다중 조건
- Filter 후 Select
- Update 후 재조회
- Delete 후 재조회
- Inner Join
- Left Join
- Right Join
- Full Join
- Left Anti / Right Anti
- Left Semi / Right Semi
- Join key no match
- Join duplicate multiplication
- Join different key column names
- GroupBy 기본 동작
- OrderBy 이후 인덱스 상태
- 정렬 없는 SelectTable / Top 결과 순서
```

### 7.2 인덱스 일관성 테스트

반드시 포함:

```text
- 인덱스 생성 후 indexed column update
- update-only upsert
- insert 발생 upsert
- AddColumn / RemoveColumn
- Compute / CopyColumn
- Value Let
- UpdateBatch
- DeleteBatch
- OrderBy 이후 rowID 의미 확인
- 컬럼명 캐시와 숫자 컬럼 index 캐시 동시 무효화
```

### 7.3 백엔드 동치 테스트

DictHashIndex와 CoreHashIndex가 모두 존재할 경우:

```text
Same input table
Same operations
Same output rows
Same row counts
Same join cardinality
Same group results
```

두 백엔드는 기능적으로 완전히 동일해야 한다.

### 7.4 벤치마크

최소 벤치:

```text
- Build index: N rows
- Where equality: hit / miss / high duplicate / low duplicate
- Multi-Where intersection
- Filter over candidate RowSet
- Inner Join: 1:1, 1:N, N:N
- Left Join no-match-heavy
- GroupBy with short keys
- GroupBy with long keys
```

측정 원칙:

- 1회 측정만으로 결론 내리지 않는다.
- warm-up 후 여러 회 측정.
- median 또는 best-of-N / median-of-N 기록.
- 기존 baseline 대비 비교 로그 유지.

---

## 8. 금지 사항과 주의 사항

### 8.1 금지 사항

아래 패턴이 DB 실행 경로에 새로 생기면 안 된다.

```vba
someDict.Keys
someDict(key)
Set rows = New Collection
Set candidates = New clsSet
candidates.LoadFromArray rowIds
keys = someSet.KeysArray
```

예외:

- DictHashIndex 내부 구현.
- 테스트 코드.
- public 호환 shim.
- 인덱스 무효화 helper 내부 metadata lookup.

### 8.2 조기 Core 전면 도입 금지

CoreHashIndex는 `DbHashIndex` 표면이 안정된 뒤 도입한다.

이유:

- 현재 병목에는 Dictionary 한계와 실행계 중간표현 낭비가 섞여 있다.
- 먼저 중간표현을 정리해야 Core의 실제 이득을 측정할 수 있다.

### 8.3 Core에 DB 의미론 주입 금지

금지 예:

```text
Core_JoinTables
Core_WhereEquals
Core_GroupBy
Core_Row
Core_TableIndex
```

허용 예:

```text
Core_clsLongList
Core_clsLongSet
Core_clsHashSlots
Core_clsSet
```

---

## 9. 성공 기준

리팩토링 성공 조건:

```text
- QueryEngine에서 Dictionary / Collection 직접 사용이 사라진다.
- row 후보 표현이 clsSet에서 DbRowSet으로 전환된다.
- 컬럼 인덱스 payload가 Collection에서 DbRowList로 전환된다.
- Join이 key set materialization 없이 index cursor / direct probe로 실행된다.
- public facade 동작은 기존과 호환된다.
- 기존 `SelectRows` / `SelectRowsByIndex` / `GetUniqueKeys` 호환 표면은 shim 또는 명시적 migration 경로로 보호된다.
- 인덱스 stale 관련 테스트가 통과한다.
- key 정규화 / case-insensitive 비교 / Empty / Error 처리 의미론이 백엔드 간 동일하다.
- DictHashIndex 기준 기능이 안정된다.
- CoreHashIndex는 선택 가능한 백엔드로 실험 가능하다.
```

성능 성공 조건은 별도다. 구조 전환 자체의 성공과 CoreHashIndex의 성능 우위는 분리해서 판단한다.

---

## 10. 최종 방향성

이 리팩토링은 기존 UnivDB를 폐기하는 작업이 아니다.

현재 프로토타입은 기능 구조를 검증하는 데 성공했다. 이제 해야 할 일은 프로토타입 단계에서 노출된 고수준 컨테이너 전달 방식을 DB 내부 실행 표현으로 낮추는 것이다.

최종 메시지:

> Dictionary를 없애는 것이 목표가 아니다. Dictionary를 DB 코어 언어에서 밀봉하는 것이 목표다.
>
> Core를 DB로 키우는 것이 목표가 아니다. Core는 DB를 모르는 저수준 관계 연산 커널로 유지한다.
>
> UnivDB의 내부 언어는 `HashIndex`, `RowList`, `RowSet`이 되어야 한다.
>
> 구현은 반드시 작은 검증 가능한 커밋으로 쪼갠다.
