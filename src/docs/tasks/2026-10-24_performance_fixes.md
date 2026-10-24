# 실측 기반 성능 수정 계획

기준: `06f5efe800444aca6c23fd93ab1fdd15ebba78e0`,
`experiments/perf-20261024/report.txt`의 실측. 사용자 승인: 계획 수립 후 착수.

## 1차 변경: 공개 계약을 유지하는 국소 수정

- Query.Update: clsTable 내부 rowID 일괄 쓰기. 모든 대상 열/행을 쓰기 전에 검증하고,
  열 이름을 한 번만 해석한다. 같은 실제 열의 중복 SetVal은 기존 마지막 값 우선 유지.
  변경 열별 인덱스/reader 값 버전은 한 번 무효화하고 행 세대는 유지한다.
  기존 UpdateBatch의 Collection 경로는 쿼리 엔진에 가져오지 않는다.
- FilterAny: 모든 arm이 같은 실제 열의 숫자 `=`이면 CDbl 키 집합으로 한 번 스캔.
  전체 arm 선검증, Empty/Error/Date/Null 제외, 숫자 문자열 및 Boolean의 CDbl 비교, 원본 rowID 순서,
  선행 후보 AND, 중복 값 제거를 유지한다. 그 외는 기존 실행 경로.
- 문서: 일괄 입력은 기존 Builder+Upsert를 사용할 수 있는 조건과 예제를 명시한다.
  저장소에는 테스트 외 업무용 반복 Insert 호출부가 없어 임의 치환하지 않는다.
  중복 키를 보존해야 하는 append와 Upsert는 동치가 아니다.

영향 범위: clsQuery, clsTable, TestColumnReader, TestIndexInvalidation,
TestChainingSemantics, TestBenchDb, UnivDB_Reference.md. 공개 시그니처/exports/config 변경 없음.

## 후속 변경 단위

2. 일괄 append: 중복 키/입력 순서/빈 입력/스키마/선검증/세대 증가 계약을 먼저 고정한 뒤
   명시적 API를 설계. 기존 단건 Insert의 즉시 반영 계약이나 저장 구조를 몰래 바꾸지 않는다.
   복수 컬럼 삭제는 누락 열 멱등성, 헤더 없는 테이블, 마지막 열 삭제 계약을 먼저 확인.
3. 희소 Filter와 Semi/Anti 경로: 밀봉된 컬럼 접근, 캐시 상태, 결과 순서를 보존하는
   별도 변경으로 검증. 현재 단일 인덱스 대조군은 중복 키 출력 순서가 달라 그대로 적용 불가.
4. Report 지연 실행, 전역 캐시, CSR 버퍼 수명은 보류. RowSet 확장은 우선순위를 낮춘다.

## Test delta brief

- Unique risk: 일괄 쓰기에서 다른 열/후보 행 오염, 유지 중인 인덱스/reader의 stale 누락;
  숫자 집합 경로에서 Empty=0 혼동, 날짜 참여, Boolean/숫자 문자열 누락, OR 순서 변경.
- Existing owner: TestColumnReader의 mutation 매개변수 사례,
  TestIndexInvalidation의 backend 매개변수 사례, TestChainingSemantics의 동일 열 IN 사례.
- Selected layer: 기존 VBA 테스트 안의 작은 메모리 fixture. 새 테스트 파일 없음.
- Expected cost delta: 위 기존 사례 확장; 자동 수집 사례 수 유지.
  TestBenchDb에 10만 행 Update/100-arm FilterAny 성능 기록 한 사례 추가 예정.
  6분짜리 Insert 실험은 기본 suite에 넣지 않는다.
- Evidence to retire: 이전 측정 대조군은 역사적 증거로 보존. 새 구현의 기준 수치는
  제품 TestBenchDb에서 재측정하고, 임시 최적화 클래스는 제품에 포함하지 않는다.

## 검증 순서

1. 기존 소유 테스트에 경계 사례 추가 후 현재 구현에서 계약 확인.
2. 제품 두 경로 수정, lint 및 vbaloop push 컴파일 검사.
3. 선택 실행 없는 프로젝트 정책에 따라 전체 Test_* 실행.
4. 같은 fixture의 전후 성능 비교. 절대 시간 gate는 추가하지 않는다.
5. 전후 결과, 최종 변경 범위, 잔여 작업을 이 문서에 기록.

## 1차 완료 결과

- clsTable.UpdateRows(Friend)로 Query.Update를 연결했다. 정규화한 열 이름별로
  마지막 값을 보존하고, rowID 범위 선검증 후 열별로 값을 쓴다.
  무효화는 쓰기 직전에 수행해 중간 오류에도 예전 인덱스가 유효하게 남지 않게 한다.
- FilterAny는 모든 arm의 실제 열 번호와 정규화된 `=` 연산자가 같을 때만
  CDbl 키 집합을 사용한다. 문자열 인덱스 키 의미론을 숫자 비교에 섞지 않았다.
- 기존 구현에서 계약 검증 중 IsNumeric(False)가 True임을 확인했다.
  따라서 False=0 의미를 기대값에 반영했고 성능 변경에서 그대로 유지한다.
- 공개 API 및 단건 Insert 동작은 변경하지 않았다. Builder+Upsert 예제에는
  같은 Row ID의 병합과 미설정 셀 Empty 덮어쓰기 주의사항을 명시했다.

동일 fixture의 전후 결과 (QPC, 워밍업 1회 + 3회 중앙값):

| 작업 | 이전 | 이후 | 배수 |
|---|---:|---:|---:|
| 10만 행 중 5만 행 Query.Update, warm 인덱스 | 161.2599ms | 21.5377ms | 7.49배 |
| 10만 행, 동일 숫자 열 100개 OR | 1432.9024ms | 63.2213ms | 22.66배 |

위 fixture는 기존 BuildJoinTable을 재사용한 2열 데이터다. 과거 보고서의 4열 fixture
절대 시간과 직접 비교하지 않고, 이번 같은 fixture의 전후를 비교한다.
성능 gate는 없으며 baseline 로그만 기록한다.
증거: `experiments/perf-20261024/fix-01/before-after.json`, `tests-before.json`, `tests-after.json`.

## 최종 Test delta

- Unique risk: 일괄 Update의 다중 열/별칭/후보 재사용/양 backend 인덱스 무효화,
  숫자 집합의 타입 변환과 순서. 기존 동일 열 IN/reader mutation/retained index 사례를 확장.
- Owner and layer: TestChainingSemantics, TestColumnReader, TestIndexInvalidation의
  메모리 fixture 및 TestBenchDb의 기존 QPC/fixture. 새 파일/실행 lane 없음.
- Added / extended / removed: 성능 기록 사례 1개 추가, 기존 정합성 사례 3개 확장, 삭제 0개.
- Collected-case delta: 125 -> 126 (+1, 전후 모두 같은 126개를 실행).
- 검증: 변경 전/후 각각 126 pass, 0 fail, 0 skipped. VBE compile, lint, diff --check 통과.
- Runtime delta: 같은 126개 suite의 test_start~마지막 test_end는 22.198초 -> 20.281초
  (-1.917초, 각 1회 관측; Excel 시작/주입 제외). 전체 suite 시간에는 다른 벤치 변동도
  포함되므로 위 두 연산의 중앙값과 구분한다. 장시간 Insert 사례는 추가하지 않았다.
- Heavy-resource launch delta: suite의 추가 Excel 시작 없음. 각 검증은 기존 vbaloop
  ephemeral push/test를 사용했으며, 제품 테스트에 별도 프로세스/네트워크 의존성 없음.
- Coverage-floor impact: 변경 없음. coverage 수치는 별도로 측정하지 않았다.

후속 2~4단계는 아직 구현하지 않았다. 특히 Insert의 중복 키 append와 조인의 출력 순서는
이번 변경으로 해결했다고 간주하면 안 된다. 배포용 dist는 생성하지 않았다.
