# 성능 계획 단계별 실측 (2026-10-31)

계획: `2026-10-31_perf_implementation_plan.md`. 수치는 `TestBenchDb` 8절 벤치(warm-up 1 + 3회 중앙값, ms),
`vbaloop test` 전량 실행 중 측정 (`_dbg\UnivDB.jsonl`의 `bench_db_*` 이벤트). 같은 PC, Dict 백엔드.
각 열은 해당 단계 커밋 직후 수치(누적).

| 벤치 (이벤트.필드) | 0 (before) | 1 | 2 | 3 | 4 | 5 | 6 | 7 |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| select_small.ms (27k x 34, 548행) | 53.1 | 5.7 | 5.7 | 5.7 | 5.2 | - | 6.0 | 6.2 |
| join_large.inner_ms (27k x 91k, 3천 키) | 699.2 | 700.9 | 236.3 | 236.3 | - | - | 211.8 | 206.3 |
| join_large.left_anti_ms | 524.6 | 521.6 | 258.5 | 258.5 | - | - | 199.3 | 194.9 |
| upsert.all_update_ms (27k x 34) | 523.7 | 518.5 | 504.6 | 152.9 | 153.1 | - | 153.3 | 136.1 |
| query_insert.one_row_ms | 439.3 | 438.3 | 472.6 | 417.9 | 408.4 | 124.8 | 139.2 | 133.0 |
| index_build_large.unique_ms (91k) | 483.7 | 457.1 | 457.1 | 457.1 | 454.1 | - | 221.5 | 215.9 |
| index_build_large.group50_ms (91k) | 165.9 | 147.6 | 148.7 | 148.7 | 146.1 | - | 122.2 | 121.2 |
| load.sheet_ms (27k x 34, Value2 읽기 포함) | - | - | - | 240.8 | 155.2 | - | 174.8 | 158.7 |
| load.array_ms (LoadFromArray) | - | - | - | 189.7 | 160.2 | - | 166.1 | 160.9 |
| load.sheet_notrim_ms (bTrim:=False) | - | - | - | - | - | - | - | 76.3 |
| load.array_notrim_ms (bTrim:=False) | - | - | - | - | - | - | - | 80.3 |

## 단계 기록

- 단계 0: 벤치 5종(`Test_BenchDb_Plan*`) + Join 행 순서 고정 테스트(`TestJoinOrder.bas`, 양 백엔드) 추가. 137 pass.
- 단계 1: `clsTable.GatherRows`/`KeepRows`(Friend) 추가. `clsQuery.SelectTable`/`Delete`, `clsTable.AppendFrom`의
  `RawArray` 전체 사본 제거, `clsTableBuilder.ExtractRow`는 `GetRecord`(1행 사본)로. `clsReportBuilder`의 clone은
  의미상 전체 사본이 필요해 유지. 140 pass. (인덱스 빌드 변동은 측정 잡음 범위)
- 단계 2: `modEngine.JoinTable`을 오른쪽 커서 탐침 1패스 + 행 쌍 목록(공통/왼쪽 전용/오른쪽 전용) + 양쪽 `GatherRows`로
  재구성. 크기 계산 패스와 양쪽 `RawArray` 제거. 왼쪽 전용은 공통 패스의 일치 플래그로 판정(재탐침 없음; left anti만
  오른쪽 탐침). `DbIndexCursor.InitCsr` 키 인자 ByRef. 행 순서 불변(TestJoinOrder green). 기존 벤치: 5k 1:1 26 -> 14,
  no-match left 34 -> 11, wide 20k x 30 inner 198 -> 130. 140 pass.
- 단계 3: `UpsertFrom`을 열 매핑 Long 배열 + 1패스(대상 행 조회 1회 + 대상 행 스탬프로 중복 검사, 삽입 키끼리만
  Dictionary) + 2패스(조회 없이 쓰기)로 재작성. 쓰기는 소스의 Friend `CopyCellsInto`, 키 열은 `CopyColumnValues`
  (1열 사본) - `tSource.RawArray` 제거. 오류 메시지/fail-fast/세대 동작 불변. `TestUpsertFrom.bas` 추가. 143 pass.
- 단계 4: `LoadFromSheet`가 Private ByRef `LoadArrayCore`를 호출(공개 `LoadFromArray`의 ByVal 복사 1회 생략,
  `LoadFromArray`도 같은 헬퍼에 위임 - 시그니처 불변). 적재 벤치 `Test_BenchDb_PlanLoad` 추가(단계 3 열에 before 기록).
  복사 제거만으로 sheet 240.8 -> 179.7, 이어서 Trim 루프를 열 우선(메모리 순서)으로 바꿔 155.2 / array 184 -> 160. 144 pass.
  (단계 2~3 열의 join `-`는 이 단계에서 영향 없음 - 재측정 생략 표기)
- 단계 5: `Query.Insert`가 새 행 1개만 매핑해 Friend `AppendRowValues`로 추가(1회 결합 복사, 새 행 문자열 셀만
  `Trim$`, 세대 증가). 계약 차이: 기존 행은 더는 재트림되지 않음(이전엔 Insert마다 `LoadFromArray`로 전 셀 재트림 +
  중복 헤더 재검사 - `AdoptArray` 결과 테이블의 동명 열에서 Insert가 실패하던 부작용도 사라짐). 남은 비용은 결합 배열
  -> `m_vData` 대입의 deep copy(문자열 셀 재할당)와 해제 - 함수 반환 대입으로 바꿔도 117ms로 차이 없어 단순형 유지.
  `TestQueryInsert.bas` 추가. 146 pass.
- 단계 6 (계약 변경): `modIndexKey.FoldIndexKey`(하이브리드 접기, Declare 모듈 내장). `DictHashIndex`는 접기 키 +
  binary Dictionary + 첫 등장 표기 배열 `mKeyText`(Cursor 키), 새 키는 Item 읽기 자동 추가 + 쓰기.
  `DictHashIndex`/`CoreHashIndex`에 Friend `BuildFromArray` - `GetHashIndexByIndex`가 구체 타입으로 분기해 `m_vData`를
  직접 넘김(열 사본 + 행당 인터페이스 `AddRow` 제거). GroupBy 복합 키는 **파트별** 접기(접힌 길이로 접두 - 전체 문자열을
  한 번에 접으면 다른 파트의 가나 유무에 따라 같은 값이 다르게 접힘), Upsert 삽입 키 Dictionary도 접기 + binary.
  Core는 의미 불변(빌드만 일괄화, 20k 빌드 80 -> 78 수준). `TestKeyFold.bas`(검증 문서 2.2절 표, 표기 보존, 판별 경계).
  Reference 6.1절/Upsert/Insert 절, CLAUDE.md/AGENTS.md 하류 벤더링 메모 갱신. 149 pass.
  20k 벤치(Dict): 인덱스 빌드 저중복 49 -> 41, 고중복 34 -> 31. 기대치(91k 250~300)보다 좋은 221ms.
  (load 열 175/166은 단계 4의 155/160과 같은 코드 경로 - 회차 편차)
- 단계 7 (공개 API 추가): `LoadFromArray(..., Optional bTrim = True)`, `LoadFromSheet(..., Optional bTrim = True)`,
  `UnivDB.RegisterTable(..., Optional bTrim = True)`. 중복 헤더 거부/헤더 Trim은 bTrim과 무관.
  `RegisterTable` 배열 경로는 자기 ByVal 사본을 Friend `LoadOwnedArray`로 넘겨 복사 1회 추가 절감.
  Reference 5절 "적재 Trim 계약" 명문화 - 계획/검증 문서의 `Filter "="`는 실제로는 문자열 연산자 `IS`/`LIKE` 경로
  (`=`는 숫자 전용, 문자열 값이면 9999). `TestLoadTrim.bas` 추가. 152 pass.

## 요약 (단계 0 -> 7, ms)

| 경로 | before | after | 배수 |
|---|---:|---:|---:|
| 소량 Select (27k x 34, 548행) | 53.1 | 6.2 | 8.6x |
| Join inner (27k x 91k) | 699.2 | 206.3 | 3.4x |
| Join left anti | 524.6 | 194.9 | 2.7x |
| UpsertFrom 전 행 갱신 | 523.7 | 136.1 | 3.8x |
| Query.Insert 1행 | 439.3 | 133.0 | 3.3x |
| 인덱스 빌드 91k 고유 | 483.7 | 215.9 | 2.2x |
| 인덱스 빌드 91k 50그룹 | 165.9 | 121.2 | 1.4x |
| 시트 적재 27k x 34 | 240.8 | 158.7 (Trim 생략 76.3) | 1.5x (3.2x) |
