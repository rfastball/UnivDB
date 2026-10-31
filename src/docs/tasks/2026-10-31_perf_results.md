# 성능 계획 단계별 실측 (2026-10-31)

계획: `2026-10-31_perf_implementation_plan.md`. 수치는 `TestBenchDb` 8절 벤치(warm-up 1 + 3회 중앙값, ms),
`vbaloop test` 전량 실행 중 측정 (`_dbg\UnivDB.jsonl`의 `bench_db_*` 이벤트). 같은 PC, Dict 백엔드.
각 열은 해당 단계 커밋 직후 수치(누적).

| 벤치 (이벤트.필드) | 0 (before) | 1 | 2 | 3 |
|---|---:|---:|---:|---:|
| select_small.ms (27k x 34, 548행) | 53.1 | 5.7 | 5.7 | 5.7 |
| join_large.inner_ms (27k x 91k, 3천 키) | 699.2 | 700.9 | 236.3 | 236.3 |
| join_large.left_anti_ms | 524.6 | 521.6 | 258.5 | 258.5 |
| upsert.all_update_ms (27k x 34) | 523.7 | 518.5 | 504.6 | 152.9 |
| query_insert.one_row_ms | 439.3 | 438.3 | 472.6 | 417.9 |
| index_build_large.unique_ms (91k) | 483.7 | 457.1 | 457.1 | 457.1 |
| index_build_large.group50_ms (91k) | 165.9 | 147.6 | 148.7 | 148.7 |

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
