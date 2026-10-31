# 성능 계획 단계별 실측 (2026-10-31)

계획: `2026-10-31_perf_implementation_plan.md`. 수치는 `TestBenchDb` 8절 벤치(warm-up 1 + 3회 중앙값, ms),
`vbaloop test` 전량 실행 중 측정 (`_dbg\UnivDB.jsonl`의 `bench_db_*` 이벤트). 같은 PC, Dict 백엔드.
각 열은 해당 단계 커밋 직후 수치(누적).

| 벤치 (이벤트.필드) | 0 (before) | 1 | 2 | 3 | 4 | 5 |
|---|---:|---:|---:|---:|---:|---:|
| select_small.ms (27k x 34, 548행) | 53.1 | 5.7 | 5.7 | 5.7 | 5.2 | - |
| join_large.inner_ms (27k x 91k, 3천 키) | 699.2 | 700.9 | 236.3 | 236.3 | - | - |
| join_large.left_anti_ms | 524.6 | 521.6 | 258.5 | 258.5 | - | - |
| upsert.all_update_ms (27k x 34) | 523.7 | 518.5 | 504.6 | 152.9 | 153.1 | - |
| query_insert.one_row_ms | 439.3 | 438.3 | 472.6 | 417.9 | 408.4 | 124.8 |
| index_build_large.unique_ms (91k) | 483.7 | 457.1 | 457.1 | 457.1 | 454.1 | - |
| index_build_large.group50_ms (91k) | 165.9 | 147.6 | 148.7 | 148.7 | 146.1 | - |
| load.sheet_ms (27k x 34, Value2 읽기 포함) | - | - | - | 240.8 | 155.2 | - |
| load.array_ms (LoadFromArray) | - | - | - | 189.7 | 160.2 | - |

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
