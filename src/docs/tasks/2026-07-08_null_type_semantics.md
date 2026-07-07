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
| Error (#N/A 등) | 행 skip | 제외 (IsNumeric=False) | **CStr 크래시 의심** | 행 skip (R1 이후) | fallback CStr **크래시 의심** |
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

우선순위: D0(크래시 실증) > D1 > D2 > D3. 각 단계 green + vh check 후 커밋.

## 6. 스코프 밖 (명시)

- 컬럼 타입 선언 시스템 (D-c) - 1.0 이후 재평가.
- 3치 논리 (D-b) - 도입하지 않기로 방향 승인.
- public shim 시그니처 변경 - 금지 사항 유지.
