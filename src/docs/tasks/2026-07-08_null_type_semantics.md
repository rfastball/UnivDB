# UnivDB 비결정 D 실행 계획: NULL / 타입 의미론

작성일: 2026-07-08
상위 문서: `../UnivDB_design_decisions_202607.md` (4장 비결정 D - "1.0 공개 전 반드시 확정"),
`2026-07-05_next_direction.md` (R1에서 하위 항목 D2(대소문자)만 선행 해소됨)
상태: **방향 승인됨 (2026-07-08)** - 착수 대기

---

## 1. 배경: 왜 D가 1.0의 관문인가

이 프로젝트의 판단 기준은 "조용히 틀린 결과가 최악의 실패 양식"이다.
현재 Empty / Error / 타입 처리가 실행 경로마다 다른 규칙을 타고 있어,
같은 데이터가 경로에 따라 다르게 집계·필터링되는 상태다. R1(D2)로 대소문자는
통일했지만 그것은 D의 하위 항목 하나였다.

## 2. 현재 실동작 매트릭스 (2026-07-07 코드 기준 조사)

| 값 | 인덱스 키 (Where/Join/GroupBy 키) | Filter 숫자 경로 | Filter 문자 경로 (IS/LIKE) | 집계 (SUM/AVG/COUNT...) | OrderBy |
|---|---|---|---|---|---|
| Empty (빈 셀) | `""` 키로 참여 | **0으로 참여** (IsNumeric(Empty)=True) | `""`로 참여 | **제외** (NULL처럼) | 숫자열에선 0 취급 |
| Error (#N/A 등) | 행 skip | 제외 (IsNumeric=False) | **locale 문자열로 조용히 참여** (D0 반증) | 행 skip (R1 이후) | **텍스트로 맨 뒤 정렬** (D0 반증) |
| 숫자 42 vs 문자 "42" | **같은 키** (CStr 병합) | 둘 다 참여 (CDbl) | CStr로 병합 | 둘 다 숫자로 합산 | 문자 우선 규칙으로 분리 비교 |
| 날짜 | CStr = **locale 의존 키** | **제외** (IsNumeric(Date)=False) | CStr 비교 | SUM 불참 | 문자 비교로 fallback |

Empty 하나가 경로에 따라 세 얼굴(NULL / 0 / 빈 문자열)을 가진다.

## 3. 위험 지점 (조용히 틀린 결과 후보)

1. **Empty의 숫자 필터 참여**: `Filter("Score","<",100)`에 빈 셀이 0으로 걸림.
   집계에선 같은 셀이 제외되므로 필터 행 수와 COUNT가 어긋나는 시나리오가 즉시 성립.
2. **AVG의 비숫자 희석**: Count는 non-Empty면 증가, Sum은 IsNumeric일 때만 가산.
   숫자 컬럼에 문자열이 섞이면 분모만 커져 평균이 조용히 낮아진다.
3. **Error 셀 크래시 의심 2곳** (착수 시 실증 1순위):
   - Filter 문자 경로 - 셀 값에 raw `CStr` 호출 (clsQuery.Filter의 IS/LIKE 분기)
   - OrderBy 비교기 fallback - `CompareValues`의 else 분기 raw `CStr`
   vbError Variant에 CStr은 type mismatch 가능성이 있다.
4. **Trim 비대칭**: Filter 문자 경로에서 비교값은 `FastNorm`(Trim 포함), 셀 쪽은 raw
   `CStr`. 셀 앞뒤 공백 시 IS 매칭 실패.
5. **날짜 키의 locale 의존**: 인덱스 키가 `CStr(date)`라 시스템 로캘에 따라 키가 달라짐.
6. **`42` == `"42"` 키 병합**: TestHashIndex가 의도된 계약으로 고정했으나 공개 문서 미기재.

## 4. 결정 축과 권장안 (2026-07-08 방향 승인)

- **D-a. Empty의 의미**: **NULL 계약 채택 권장** - 모든 비교·집계 불참, 인덱스 키만 `""` 유지.
  집계는 이미 NULL식이라 실변경은 숫자 필터의 0 참여 제거 하나. 변경 폭 최소로 일관성 확보.
- **D-b. 3치 논리(SQL NULL) 전면 도입**: **비도입 권장**. Excel 사용자 기대(빈 셀=빈 값)와
  어긋나고 구현 비용 대비 이득 없음. "Empty는 항상 불참" 단순 규칙으로 충분.
- **D-c. 컬럼 타입 선언(스키마)**: **비도입 권장 (이월)**. 입출구가 Range.Value2 = 2D Variant인
  이상 선언은 검증 레이어일 뿐이며 1.0 스코프를 키운다.
- **D-d. 키 계약 명문화**: R1~R3으로 사실상 확정된 "정규화 문자열 키" 계약(CStr 병합,
  Empty==`""`, Error skip, 대소문자 무시, 길이 접두 복합 키)을 `UnivDB_Reference.md`에 공개
  문서화. 날짜 canonical 포맷(예: ISO)을 키 정규화에 포함할지는 이 단계에서 함께 결정.
- 부수 항목: GroupBy 그룹 값 String화(타입 보존 여부), AVG 분모 규칙, Trim 대칭화.

## 5. 실행 계획 (R1 검증된 2단 커밋 패턴)

검증 도구는 `vbaflow`다 (2026-07-08 `_VbaHarness`/`vh`에서 전환). 각 Phase 루프:
`vbaflow session start` → src 편집 → `vbaflow push` → `vbaflow test` (green =
`payload.summary.fail == 0`) → `git commit`. 커밋 전 규약 점검은 `vbaflow lint`
(구 `vh check`). 실패 어서션은 `vbaflow test --json`의 `payload.tests[].detail`에 인라인.
모든 명령 `--json` 필수. 상세 절차는 `.claude/skills/vbaflow/SKILL.md`.

### Phase D0 - 현행 고정 회귀 + 크래시 실증
- 매트릭스의 각 칸을 드러내는 회귀 테스트 추가 (TestNullTypeSemantics.bas 신설).
- Error 셀 크래시 의심 2곳(위 3-3)의 실동작을 여기서 확인 - 크래시면 즉시 fail-fast 또는
  skip 계약으로 별도 fix 커밋.
- 커밋: `test: pin current null and type semantics`

### Phase D1 - Empty NULL 계약 적용 (D-a)
- Filter 숫자 경로에서 Empty 불참 처리 (IsEmpty 가드).
- 기대값 변경은 별도 커밋: `feat: exclude empty cells from filter comparisons`

### Phase D2 - 집계·필터 정합 (부수 항목)
- AVG 분모를 "숫자로 가산된 셀 수"로 교정, Filter 문자 경로 Trim 대칭화.
- 커밋: `fix: align aggregate denominators and filter trim symmetry`

### Phase D3 - 키 계약 명문화 (D-d)
- `UnivDB_Reference.md`에 키 의미론 장 신설. 날짜 canonical 포맷 결정 포함.
- 커밋: `docs: publish key semantics contract`

우선순위: D0(크래시 실증) > D1 > D2 > D3. 각 단계 green(`vbaflow test`) + `vbaflow lint` 후 커밋.

주의: D0 크래시 실증 대상(Filter 문자 경로, OrderBy 비교기)은 전부 dict 경로라 현재
green이지만, Core 백엔드 컴파일 결함(CLAUDE.md 알려진 미결)이 있어 Core 경로 테스트는
모달을 띄운다. D 작업의 신규 테스트는 dict 경로만 다루므로 영향권 밖 - 단, 새 테스트가
실수로 `IndexBackend = dbBackendCore`를 건드리지 않도록 주의.

## 5.5. D0 실행 결과 (2026-07-10)

`TestNullTypeSemantics.bas`로 매트릭스 13칸을 실행 고정. 11칸은 조사값 그대로 확정.
**크래시 의심 2곳(3-3)은 반증됨** - 이 환경의 `CStr(CVErr(xlErrNA))`는 크래시 없이
locale 문자열(`"오류 2042"`)을 반환한다 (errno=0). 실제 귀결:

- **Filter 문자 경로(IS)**: Error 셀이 `"오류 2042"`로 문자열화돼 **조용히 비교에 참여**한다.
  `IS "a"`엔 불매칭(1행), `IS NOT "A"`엔 Error 행이 남는다(2행).
- **OrderBy**: 크래시 없이 Error 셀이 **텍스트로 맨 뒤 정렬**(숫자는 숫자끼리 정상, 한/영 로캘 공통).

즉 실패 양식은 "크래시"가 아니라 **"조용히 틀린 결과"** - 프로젝트 철학상 더 나쁜 쪽.
따라서 D0에 즉시 크래시 fix는 없고, **Error-skip 통일은 D1급 의도적 계약 변경**으로 이월한다
(Empty NULL 계약과 같은 원칙: "Empty=NULL 불참", "Error=skip 어디서도 불참").
핀은 locale 문자열을 하드코딩하지 않고 locale-무관 불변식(errno=0, 참여 여부, 정렬 보존)으로 고정.

### D1 확장 (D0 결과 반영)
- Empty NULL 계약(D-a: 숫자 필터 0 참여 제거)에 **Error-skip 통일**을 합류시킨다:
  - Filter 문자 경로(`clsQuery` IS/IS NOT/LIKE/NOT LIKE): Error 셀 비교 제외.
  - OrderBy 비교기(`clsTable.CompareValues`): Error 셀 안전 취급. 이미 존재하는
    `Core_modCompare.Core_CompareValues`(Error/Date/Numeric 인지, `Core_SafeText` 경유)로
    배선하는 안이 유력 - 단 날짜 정렬을 locale 텍스트→숫자로 바꾸므로 그 변경은 별도 판단(D2/D3).
- 회귀 펜스: `nullsem_err_strfilter_isnot_a`(2→1), OrderBy 행 보존/정렬 위치.

## 5.6. D1 실행 결과 (2026-07-10 완료)

인덱스를 안 거치는 직접 비교 표면 3곳을 정합화. 커밋 3개:
- `4ef7ba7` **MIN/MAX pin** (신규 표면, Explore 발견): GroupBy MIN/MAX는 셀에 raw Variant `>`/`<`.
  Error 셀은 미가드라 compare에 닿으면 **Type Mismatch(13) 크래시**로 실증됨(fix는 D2).
- `67b7900` **feat: Empty=NULL** (숫자 필터): `Not IsEmpty` 가드로 Empty의 0 참여 차단.
  핀 flip: `numfilter_lt100` 3→2, `eq0` 1→0, `filter_vs_count` 3→2 (필터-COUNT 정합).
- `62870f9` **fix: Error=skip/safe** (문자 필터 + OrderBy): `clsQuery` 문자 경로 4연산에
  `Not IsError` 가드(부정 연산 포함 skip), `clsTable.CompareValues`에 Error last 결정적 배치.
  핀 flip: `strfilter_isnot_a` 2→1. OrderBy 핀은 불변(기계적 하드닝, locale-무관 보장).

확정된 계약: **Empty=NULL(숫자 비교 불참, 문자/키의 ""는 유지) + Error=skip(비교/정렬 표면
어디서도 값 참여 없음)**. 인덱스 경로는 이미 정합이라 무접촉. Core_CompareValues는 dead code라
D1에서 배선 안 함(날짜 의미론까지 바꿔 범위 초과) - 통합 판단은 D2/D3.

## 5.7. D2 실행 결과 (2026-07-10 완료)

집계·필터 정합. 커밋 2개 (MIN/MAX fix는 집계 계약에 합류):
- `f384efe` **fix: 집계 Error=skip + AVG 희석 교정** (단일키/복합키 두 경로): COUNT는 Error
  제외(D1 일관), AVG 분모를 "숫자로 가산된 셀 수"로 교정(함수별 count 정책 분리), MIN/MAX는
  `Not IsError` 가드로 raw Variant 비교 크래시 차단. 핀 flip: `err_agg_count` 3→2,
  `err_agg_avg` 40/3→20, `date_agg_avg` 5→10, `minmax` errno 13→0(+MAX 30/MIN 10).
- `6bca416` **fix: 문자 필터 Trim 대칭화** (위험 지점 4): 셀 쪽 `CStr`→`FastNorm`으로 target과
  대칭 정규화. 셀 앞뒤 공백 있어도 IS 매칭 성립. 신규 핀 `strfilter_trim_is_a`.

집계 계약 확정: **Empty=NULL·Error=skip 둘 다 COUNT 포함 모든 집계 불참, 비숫자 실값(Date/문자)은
COUNT 포함·SUM 불참·AVG 분모 제외**. Empty=NULL과 Error=skip이 비교/정렬/집계/필터 전 표면에서 통일됨.

## 5.8. 남은 작업 (D3)

- **날짜 의미론**: OrderBy 날짜 정렬 text→numeric, 인덱스 날짜 키 canonical(ISO) 포맷,
  dead code `Core_CompareValues` 통합 여부. (미결 - 방향 미정)
- **키 계약 명문화 (D3)**: `UnivDB_Reference.md`에 키 의미론 장 신설(CStr 병합, Empty=="",
  Error skip, 대소문자 무시, 길이 접두 복합 키, + 위 NULL/타입 집계 계약).

## 6. 스코프 밖 (명시)

- 컬럼 타입 선언 시스템 (D-c) - 1.0 이후 재평가.
- 3치 논리 (D-b) - 도입하지 않기로 방향 승인.
- public shim 시그니처 변경 - 금지 사항 유지.
