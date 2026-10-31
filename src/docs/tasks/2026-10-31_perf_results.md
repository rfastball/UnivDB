# 성능 계획 단계별 실측 (2026-10-31)

계획: `2026-10-31_perf_implementation_plan.md`. 수치는 `TestBenchDb` 8절 벤치(warm-up 1 + 3회 중앙값, ms),
`vbaloop test` 전량 실행 중 측정 (`_dbg\UnivDB.jsonl`의 `bench_db_*` 이벤트). 같은 PC, Dict 백엔드.
각 열은 해당 단계 커밋 직후 수치(누적).

| 벤치 (이벤트.필드) | 0 (before) | 1 |
|---|---:|---:|
| select_small.ms (27k x 34, 548행) | 53.1 | 5.7 |
| join_large.inner_ms (27k x 91k, 3천 키) | 699.2 | 700.9 |
| join_large.left_anti_ms | 524.6 | 521.6 |
| upsert.all_update_ms (27k x 34) | 523.7 | 518.5 |
| query_insert.one_row_ms | 439.3 | 438.3 |
| index_build_large.unique_ms (91k) | 483.7 | 457.1 |
| index_build_large.group50_ms (91k) | 165.9 | 147.6 |

## 단계 기록

- 단계 0: 벤치 5종(`Test_BenchDb_Plan*`) + Join 행 순서 고정 테스트(`TestJoinOrder.bas`, 양 백엔드) 추가. 137 pass.
- 단계 1: `clsTable.GatherRows`/`KeepRows`(Friend) 추가. `clsQuery.SelectTable`/`Delete`, `clsTable.AppendFrom`의
  `RawArray` 전체 사본 제거, `clsTableBuilder.ExtractRow`는 `GetRecord`(1행 사본)로. `clsReportBuilder`의 clone은
  의미상 전체 사본이 필요해 유지. 140 pass. (인덱스 빌드 변동은 측정 잡음 범위)
