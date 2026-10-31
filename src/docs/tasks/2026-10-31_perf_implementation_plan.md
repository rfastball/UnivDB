# 성능 개선 구현 계획 (2026-10-31)

새 세션 시작용 단독 문서. 근거와 실측은 `2026-10-31_perf_brief_verification.md`
(측정 모듈/원시 결과 `experiments/perf-20261031/`). 소비자 원 브리프는
`..\지출 리팩토링\src\docs\tasks\perf-synth\univdb-perf-brief.md` - 일부 수치/주장이 틀렸으니 검증 문서를 우선한다.
아래 줄 번호는 커밋 97de97e 기준 근사치.

## 0. 확정된 결정

1. Join 결과 행 순서 유지 (작은 쪽 구동 최적화는 하지 않음).
2. 하이브리드 키 접기의 계약 차이 수용: 반각 가나 탁음 결합(`ガ`/`ｶﾞ`)은 같은 키로(StrComp 쪽으로 정렬),
   현행 Dict가 미할당/서식 제어 문자끼리 잘못 합치던 것은 분리. Core 백엔드는 그대로 - 이 문자들에서만 Dict/Core가 다름.
3. `clsSet`은 범위 밖 (Variant 키 의미론이 다르고 핫 경로 아님).
4. 릴리스는 7단계 완료 후 2.10.0 한 번. 설치본(`%APPDATA%\Microsoft\AddIns\UnivDB.xlam`) 교체는 사용자 확인 후에만.
5. Core 기본 백엔드화, `UCase$`만의 키 접기, `StrConv` 접기, `AdoptArray` 무복사 이동, Trim 루프 "변경 시만 재할당"은 기각(실측).

## 1. 공통 진행 규칙

- 단계마다: src 편집 -> `vbaloop push` -> `vbaloop test` green -> 벤치 전후 수치 기록 -> 커밋. green 전 커밋 금지.
- 단언은 `LogShim.AssertEq`. 매크로 호출은 모듈 한정. src는 UTF-8 무 BOM + CRLF + 끝 개행 1개, CP949 밖 문자 금지.
- 세션 Excel 직접 COM 접근 금지(확인은 `vbaloop diff`).
- 공개 시그니처 불변(1~6단계). 새 공개 표면은 7단계(Optional 인자)와 8단계뿐.
- Public 프로시저에 `@desc` 규약 준수(상위 CLAUDE.md 2절). Friend/Private에는 넣지 않음.
- 측정 데이터 형태: 마스터형 27,401 x 34, DW형 91,000 x 27, 셀 2/3 문자열 (`experiments/perf-20261031/ZPerfProbe2.bas`의 `MakeData` 참고).

## 2. 단계

### 단계 0 - 기준선과 안전망 (커밋 1개)

- `TestBenchDb.bas`에 벤치 추가 (warm-up 1 + 3회 중앙값, `LogShim.LogEvent "bench_db_*"`):
  `select_small`(인덱스 warm, 결과 약 550행), `join_inner`(27k x 91k, 일치 3천 키), `join_left_anti`,
  `upsert_all_update`(27k x 34), `query_insert_one`, `index_build_91k`(고유/50그룹).
- **Join 행 순서 고정 테스트를 현행 코드로 먼저 작성**: inner/left/full/left anti/right anti/semi 각각, 중복 키(1:N, N:N)와
  양쪽 전용 키가 섞인 작은 테이블로 결과 전체 행 순서를 단언. 2단계 리팩토링의 회귀 그물.
- 현행 수치를 "before"로 jsonl에 남김.

### 단계 1 - 내부 수집 메서드 `GatherRows` (커밋 1개)

- `clsTable`에 추가:
  `Friend Sub GatherRows(ByRef rowIds() As Long, ByVal n As Long, ByRef srcCols() As Long, ByRef dest As Variant, ByVal destRow As Long, ByVal destCol As Long)`
  - `rowIds(0..n-1)`의 각 행을 `srcCols` 순서로 `dest(destRow + i, destCol + k)`에 복사. rowId = 0이면 Empty 유지(outer 조인용).
  - 자기 `m_vData`를 직접 읽음 - `RawArray` 전체 복사 제거가 목적. 세대 검증은 진입점 책임(Friend, 호출자 신뢰).
- 적용:
  - `clsQuery.SelectTable` (clsQuery.cls 약 672행 `vRawSource = m_SourceTable.RawArray`) - 행 목록이 이미 있음.
  - `clsTable.AppendFrom` (약 633행 `incoming = source.RawArray`).
  - `clsQuery.Delete` (약 535행): 남길 행 목록을 만들어 `clsTable`의 새 Friend `KeepRows(rowIds, n)`이 자기 배열을 재작성(세대 증가 포함).
  - 그 외 `RawArray` 내부 사용처(`clsTableBuilder` 약 129행, `clsReportBuilder` 약 54행 clone)는 의미 확인 후 필요 시만.
- 기대: 작은 Select 58.5 -> 수 ms (실측: 96%가 원본 복사).
- 테스트: GatherRows의 rowId 0/열 재배치/헤더 행 처리 단위 테스트 + 기존 Query/Delete/Append 테스트 green.

### 단계 2 - Join 행 쌍 목록 (커밋 1개)

- `modEngine.JoinTable` (약 29~268행) 재구성:
  - 현행: 크기 계산 패스(약 79~117행)와 채우기 패스(약 175행~)가 각각 커서를 돌며 상대 인덱스를 탐침(91k 키 1회전 210ms x 2).
  - 변경: 탐침 1패스에서 (왼쪽 rowId, 오른쪽 rowId) 쌍을 증가형 Long 배열 2개에 기록. 순서는 현행과 동일하게
    공통 키(오른쪽 커서 순, 키 내 a x b 중첩 순) -> 왼쪽 전용 -> 오른쪽 전용. semi/anti는 한쪽만 기록.
  - 쌍 개수 = 결과 행 수 -> 결과 배열 1회 할당 -> 왼쪽/오른쪽 `GatherRows` 각 1회. 양쪽 `RawArray`(약 133~134행) 제거.
  - 헤더 충돌 처리 코드는 그대로.
- 덤: `DictHashIndex.Cursor()`가 키 배열을 3회 복사(`mDict.Keys` -> `InitCsr(ByVal keys)` -> `mKeys = keys`) - `InitCsr`를 ByRef로.
- 기대: inner 714 -> 약 270ms, left anti 580 -> 대폭 감소.
- 테스트: 단계 0의 순서 고정 테스트 + `TestJoin`/`TestJoinDirect` green.

### 단계 3 - UpsertFrom (커밋 1개)

- `clsTable.UpsertFrom` (약 1090행):
  - 열 매핑을 `srcIdx() As Long / dstIdx() As Long`으로 1회 구축(`For Each vKey In dicColMap.Keys` 제거).
  - 1패스: 소스 행마다 `TryGetFirstRowId`로 대상 행 조회 -> `tgtRow(srcRow)` 저장 ->
    스탬프 배열 `stamp(1 To 대상 행 수)`로 중복 검사(같은 대상 행 재점유 = 중복 소스 키; 서로 다른 정규화 키는 서로 다른 첫 행).
    미존재 키(삽입 대상)끼리의 중복만 작은 Dictionary로 검사. 오류 메시지는 현행과 동일(키와 두 소스 행 번호).
  - 검증은 모든 쓰기 전에 끝남(현행 fail-fast 계약 유지).
  - 2패스: `tgtRow`로 조회 없이 셀 쓰기. 삽입 행은 결합 배열 할당 후 `GatherRows`로 소스에서 직접 수집(Collection 제거).
  - `tSource.RawArray` 제거: 셀 쓰기는 소스의 Friend 메서드(예: `CopyCellsInto(ByRef dest, srcRows, dstRows, srcCols, dstCols)`)로.
  - `BumpGeneration` 동작은 이번에 바꾸지 않음(8단계에서 함께).
- 기대: 515 -> 약 300ms 이하 (seenKeys 패스 62ms 제거 + 호이스팅 약 150ms + 소스 복사 제거).
- 테스트: 기존 Upsert 테스트(06f5efe의 중복 소스 키 거부 포함) + 삽입/갱신 혼합 + Error 키 skip.

### 단계 4 - 적재 복사 1회 제거 (커밋 1개)

- `clsTable.LoadFromSheet`(약 65행)가 `LoadFromArray`(ByVal) 대신 Private ByRef 헬퍼
  (예: `LoadArrayCore(ByRef vArr, ByVal bHasHeader)`)를 호출. `LoadFromArray`도 같은 헬퍼에 위임.
- 공개 `LoadFromArray(ByVal ...)` 시그니처 불변. 중복 헤더 거부(`BuildHeaders bRejectDup=True`)와 Trim 유지.
- 기대: 시트 적재당 약 60~70ms(문자열 위주 27k x 34).

### 단계 5 - Query.Insert (커밋 1개)

- `clsQuery.Insert`(약 564행)는 `RawArray` + 원소 복사 + `LoadFromArray`(복사 2회 + 전 셀 Trim)로 1행에 490ms.
- `clsTable`에 Friend `AppendRowValues(ByRef rowVals)`: `AppendFrom`식 1회 결합 복사, 새 행의 문자열 셀만 `Trim$`, 세대 증가.
- 컬럼 미존재 시 현행과 같은 오류. 기대: 수십 ms.

### 단계 6 - 하이브리드 키 접기 + 인덱스 일괄 빌드 (계약 변경, 커밋 1~2개)

- `modIndexKey.bas`에 `Public Function FoldIndexKey(ByRef s As String) As String` 추가:
  - 문자열에 U+3000-30FF, U+3131-318E, U+FF00-FFFF 코드 유닛이 있으면 `LCMapStringW(GetThreadLocale, LCMAP_UPPERCASE Or LCMAP_HALFWIDTH)`,
    아니면 `UCase$`. 판별은 `b() = s` 바이트 배열 상위 바이트 검사. 구현 원형: `experiments/perf-20261031/ZPerfProbe.bas`
    (`NeedsFoldBytes`, `FoldLc`).
  - Declare는 이 모듈 안에 `#If VBA7` 분기로 둔다 - UnivFilter가 벤더링하는 파일이라 Core_*에 의존하면 안 됨.
  - `NormalizeIndexKey` 계약은 불변(접기는 그 다음 단계).
- `DictHashIndex.cls`:
  - `CompareMode` 기본(binary), 키 = `FoldIndexKey(NormalizeIndexKey(key))`.
  - 원본 표기 배열 `mKeyText()`를 keyId 순서로 유지(첫 등장 표기). `Cursor()`는 이 배열을 키로 넘김 - Cursor.Key/GroupBy 렌더 표기 불변.
  - 새 키 등록은 `c0 = d.Count: id = d(k): If d.Count > c0 Then d(k) = c0 + 1` (Item 읽기 자동 추가 + 쓰기; exists+Add 대비 91k 고유에서 빠름).
- 일괄 빌드: `DictHashIndex`/`CoreHashIndex`에 `Friend Sub BuildFromArray(ByRef data As Variant, ByVal col As Long, ByVal firstRow As Long, ByVal lastRow As Long)`.
  `clsTable.GetHashIndexByIndex`(약 378행)는 `CopyColumn1D` + 행당 `AddRow`(ByVal 3~4회 + 인터페이스 디스패치) 대신
  구체 타입으로 분기해 호출(인터페이스엔 Friend를 둘 수 없음). 문자열 셀은 인라인 대입, 그 외만 `NormalizeIndexKey`, Error skip.
- GroupBy 복합 키 Dictionary(`dicAccumulator`/`dicGroupParts`, clsTable 약 756~758행)도 접기 + binary로 통일.
- UpsertFrom의 삽입 키 중복 Dictionary(단계 3)도 접기 + binary.
- 테스트(신규, Dict 백엔드): 검증 문서 2.2절 표를 그대로 단언 - 전각 라틴/숫자/소문자/공백, 반각 한글 자모, 대소문자, `é`/`É`
  같음; `e`/`é`, `ß`/`ss`, 가타카나/히라가나 다름; `ガ`/`ｶﾞ` 같음. 원본 표기 보존(첫 등장) 단언.
  기존 `Test_HashIndex_UnicodeEquality`(TestHashIndex.bas 약 212행)는 Core 전용이라 그대로.
- 문서: `UnivDB_Reference.md` 6장 "Unicode 한계" 갱신(Dict 접기 규칙, Dict/Core 차이 문자). CLAUDE.md "하류 벤더링"에
  modIndexKey 변경 = UnivFilter 재스냅샷 필요 메모.
- 기대(공개 경로): 27k 80 -> 약 65ms, 91k 498 -> 약 250~300ms, 182k 1,465 -> 약 650ms. Join 탐침도 함께 빨라짐.

### 단계 7 - Trim 생략 옵션 (공개 API 추가, 커밋 1개)

- `clsTable.LoadFromArray(vArr, Optional bHasHeader = True, Optional bTrim As Boolean = True)`,
  `UnivDB.RegisterTable(sTableName, Source, Optional bTrim As Boolean = True)` -> `LoadFromSheet`에도 전달.
- 중복 헤더 거부는 bTrim과 무관하게 유지.
- Reference에 명문화: 적재 시 문자열 셀 앞뒤 공백 `Trim$`(현재 암묵 계약), `bTrim:=False`면 `Filter "="`(비교 시 `FastNorm`이 Trim)와
  `Where`/Join/GroupBy(인덱스 키, Trim 없음) 결과가 공백 포함 셀에서 달라질 수 있음.
- 기대: 27k x 34 적재당 약 100ms.

### 단계 8 - 변경 행 보고 API (설계 확정 후 별도)

- 착수 전 결정: 비교 모드 Enum(엄격 = VarType+값 / 시트 등가 = 문서화된 느슨 규칙), "동등하면 쓰지 않음"(대상 값 유지)으로
  outChangedRowIds = 실제로 쓴 행. 소비자 `SameCellValue`(지출 리팩토링 `src/Q_Utils.bas`)의 locale 의존(`CDate`/`IsNumeric`)은 그대로 이식 금지.
- 함께: 갱신만 있는 Upsert의 `BumpGeneration`을 변경 열 무효화로(`UpdateRows`와 동일 방식), 변경 0건이면 생략.

## 3. 릴리스

1. 7단계까지 green 후 `vbaloop.toml` [dist] version = "2.10.0", `UnivDB_Reference.md` 버전 동기.
2. `TestBenchDb` before/after 표를 검증 문서 또는 새 결과 문서에 기록.
3. `vbaloop push` -> `vbaloop dist --out <path>`. 설치본 교체는 사용자 확인 후.
4. 소비자 후속(지출 리팩토링, 별도 세션): 애드인 교체와 동시에 DW 테이블의 `IndexBackend = 1` 줄 제거
   (`Q_Logic_Return_Check.bas`, `Q_Logic_Bill_Missing.bas`) - 2.9.x Core는 91k에서 약 525ms로 느려짐. `MIN_UNIVDB_VERSION` 상향,
   `X_DwIndexBackend` 재측정. UnivFilter 재스냅샷 필요 여부 판단(modIndexKey 변경).
