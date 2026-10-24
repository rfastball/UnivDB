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

1차 완료 당시 후속 2~4단계와 배포 빌드는 남아 있었다. 아래 2차 결과가 최종 상태다.

## 완주 범위와 2차 Test delta brief

사용자 승인: 완주까지 자율 진행. 아래 구현/검증/채택하지 않는 항목의 근거 기록 및
로컬 배포본 빌드까지 수행한다. 외부 게시/배포는 하지 않는다.

- AppendFrom(clsTable): 기존 스키마에 열 이름으로 매핑해 행을 한 번에 추가한다.
  중복 키/입력 순서/셀 타입 보존, self-append 지원, 기존 행 불변, 성공 시 세대 +1.
  열 이름이 완전하고 유일한 동일 스키마만 허용하며 불일치는 쓰기 전에 9999.
- RemoveColumn/RemoveColumns: 한 내부 구현으로 합쳐 배열을 한 번만 재작성.
  누락 이름은 무시하고 인자를 순서대로 해석한다. 전 열 삭제는 9999로 선거부해
  기존 부분 삭제를 막는다. 헤더 없는 테이블도 COLUMNn 이름 재계산을 지원한다.
  변경한 batch당 세대 +1, 실제 삭제가 없으면 세대 유지.
- Filter/FilterAny: 내부 후보 컬럼 복사를 사용해 반환 배열을 후보 위치로 조밀하게 만든다.
  공개 reader의 snapshot/세대 계약은 변경하지 않고, 임계값/전역 캐시도 추가하지 않는다.
- Join: 기존 양쪽 인덱스와 키별 출력 순서를 유지한다. Semi/Anti에서 미출력 입력의
  RawArray 복사만 제거하고 RightSemi의 불필요한 왼쪽 rowID span 복사도 생략한다.
- RowSet 용량/CSR/Report lazy/자동 단일 인덱스/조기 투영은 추가 구현하지 않는다.
  앞선 실측 대비 비용/호환성 근거로 채택 여부를 최종 표에 명시한다.

Test delta brief:
- Unique risk: append 중복 키 소실/열 오매핑/오류 후 오염/세대 누락, batch drop의
  마지막 열과 헤더 없는 데이터 손상, 후보 위치와 rowID 혼동, 조인 순서/우측 prefix 손상.
- Existing owner: TestBaseline에 append 계약 2개, TestChainingSemantics의 컬럼 삭제,
  TestColumnReader의 값 패리티, TestJoinDirect의 join type 사례 확장.
- Selected layer: 같은 VBA 메모리 fixture 및 기존 TestBenchDb.
- Expected cost delta: 계약 사례 +2, 짧은 append/drop 벤치 +1; 기존 사례 확장.
  장시간 반복 Insert 재측정은 생략하고 기존 362.9초 실증값을 참조한다.
- Evidence to retire: 없음. 기존 generic join/reader 경로는 계속 필요한 공개 계약이다.

## 2차 완료 결과 (v2.9.0)

- AppendFrom: 열 이름 매핑 후 전체 배열을 한 번 확장한다. 같은 키의 여러 행도 보존한다.
  예전 Insert는 그대로이며 자동 버퍼링하지 않는다. 무초기화 빈 소스는 무동작,
  헤더만 있는 소스는 스키마 검사 후 무동작이다. 기존 대상 스키마는 필수다.
- RemoveNamedColumns 내부 경로를 단수/복수/Report에서 공유한다. 단수 RemoveColumn의
  쉼표 포함 이름은 literal로 보존하고, 복수 인자만 기존 파싱 규칙을 적용한다.
  전 열 삭제 오류가 과거의 부분 삭제/오류 9에서 원자적 거부/9999로 명시화되었다.
- Filter는 CopyCandidateColumn으로 후보 위치에 대응하는 값만 복사한다.
  새 장기 snapshot/cache가 아니므로 별도 세대 객체를 만들지 않는다. 쿼리 진입점이
  후보의 세대를 검증하고, 공개 DbColumnReader 계약은 그대로 유지한다.
- Semi/Anti는 출력하지 않는 쪽의 RawArray 복사를 생략했다. RightSemi는 왼쪽
  rowID span 대신 개수만 확인한다. 양쪽 인덱스/키 순회/중복 행 순서는 유지한다.

같은 fixture의 최종 관측 (QPC 워밍업 1회 + 3회 중앙값, ms):

| 작업 | 수정 전 | 수정 후 | 해석 |
|---|---:|---:|---|
| 10만 행에서 후보 1행 Filter | 8.27 | 0.017 | 전체 컬럼 복사 제거 |
| 후보 1천 행 Filter | 8.65 | 0.36 | 24.3배 |
| 후보 1만 행 Filter | 10.95 | 3.13 | 3.5배 |
| 후보 10만 행 Filter | 35.03 | 30.69 | 밀집 후보도 퇴행 관측 없음 |
| 2만 행 x 30열에서 20열 삭제 | 323.49 | 35.86 | 9.0배 |
| 5만 행 + 1만 행, 4열 append | Insert 362887.21 | 62.68 | 기존 Insert는 1회 관측, append 준비 제외 |
| 위 append, 소스 구성까지 포함 | Builder+Upsert 345.90 | 146.12 | 신규 고유 키 fixture에서만 결과 동치 |

Insert 6분 실행은 재반복하지 않았다. Append 준비 포함/제외 시간을 모두 보존했으며
초기 대상 테이블 생성은 양쪽에서 제외했다. 준비 포함 경로는 1만 행 소스 Load와 ID 설정을
포함한다. 일반 중복 입력은 Upsert와 동치가 아니며 별도 작은 계약 테스트로 검증했다.
1차 Update/숫자 OR의 통제된 전후 결과는 위 1차 표를 참조한다. 최종 suite에서도 기록하되
실행 환경이 다른 절대 시간을 섞어 최고 성능을 주장하지 않는다.

후반 실험에서는 수정하지 않은 RawArray/SelectColumns 대조군까지 2~4배 느려졌다.
따라서 10만 x 5천 조인의 과거/현재 절대 시간만으로 개선이나 회귀를 확정하지 않았다.
추가로 `9ec684a`의 JoinTable 본문을 복제 워크북에만 넣고 동일 프로세스에서 구/신 순서를
번갈아 실행했다. 2만 행 x 30열 왼쪽, 5천 행 x 2열 오른쪽, warm 인덱스 기준:

| 조인 | 이전 | 현재 | 판단 |
|---|---:|---:|---|
| LeftSemi | 39.14 | 31.63 | 1.24배 |
| LeftAnti | 113.75 | 114.72 | 유의한 개선 주장 안 함 |
| RightSemi | 31.71 | 22.79 | 1.39배 |
| RightAnti | 36.21 | 25.83 | 1.40배 |

제품 테스트는 Dict/Core 양 backend의 중복 키 순서/우측 prefix를 확인했다.
추가 paired 실험은 Dict backend에서 모든 셀의 타입/값/순서를 비교했다.
증거: `experiments/perf-20261024/fix-02/`의 measurements.jsonl(92 assert 통과),
join-paired.jsonl(32 assert 통과), before-after.json, join-paired-summary.json,
tests-after.json, validation.json. 원래 baseline 파일은 보존했다.

### 후보별 최종 결정

| 후보 | 결정과 근거 |
|---|---|
| RowSet 배수 확장 | 보류. 10만 Add 약 4ms 절감; FromAll은 이미 사전 할당. 광범위한 자료구조 변경 대비 낮은 효과 |
| Update 일괄 쓰기 | 완료. 선검증 + 열별 무효화, 공개 Update 계약 유지 |
| 반복 Insert | 명시적 AppendFrom 완료. 단건 Insert 변경 없이 중복 보존 bulk 경로 제공 |
| Semi/Anti 단일 인덱스 | 미채택. 중복 출력 순서 변경 및 warm Semi 퇴행 실증. 불필요한 복사만 제거 |
| 컬럼 사본/희소 Filter | 후보만 복사 완료. 전역 cache/임계값/직접 public getter 루프 불필요 |
| FilterAny 동일 열 다중 조건 | 숫자 = 집합 경로 완료. IS/혼합 연산자는 기존 의미 유지 |
| Report 반복 물질화 | 지연 실행 미채택. 즉시 실행/오류 시점 계약 유지. 연속 조회는 Query 경로 활용 |
| 복수 컬럼 추가/삭제 | 삭제만 완료. AddColumns는 10만 행 8열 약 16ms로 우선도 낮음 |
| 희소 RowSet ToArray | 보류. 높은 rowID 100회 약 60ms; 별도 sparse 구조 비용 정당화 안 됨 |
| 조기 컬럼 투영 | 자동 재배치 미채택. 폭 넓은 입력에서 명시적 SelectColumns는 사용 가능 |
| RawArray/CSR | 조인 내 불필요 복사 제거만 채택. 공개 사본 계약 유지; CSR peak memory 미측정 |

### 최종 검증과 비용

- 제품: VBE compile 및 lint 49개 파일 통과. 전체 129 pass, 0 fail, 0 skipped,
  sessionDied=False. 초기 실증 125개 대비 +4, 1차 126개 대비 +3.
- 2차 추가: append 계약 2개 + 짧은 bulk 벤치 1개. 기존 컬럼 삭제/reader/조인 사례 확장.
  새 파일/lane/추가 테스트 프로세스/coverage floor 변경 없음. 자동 수집은 계속 전량 실행.
- 추가 3개 사례는 최종 실행에서 합계 0.871초(0.202 + 0.096 + 0.573).
  전체 suite 26.565초, 1차 후 20.281초보다 +6.284초이나 대조군 환경 변동도 포함한다.
  장시간 Insert 실험은 기본 suite에 포함하지 않았다. 수치 coverage는 별도 측정 안 함.
- 테스트/진단의 불필요한 중복 증거를 추가하지 않고 기존 benchmark/fixture를 재사용했다.
- 버전은 공개 API 추가에 맞춰 Reference와 vbaloop.toml 모두 2.9.0.
  로컬 dist 컴파일 결과는 fix-02/distribution.json에 보존한다. 외부 게시/배포는 범위 밖이다.
