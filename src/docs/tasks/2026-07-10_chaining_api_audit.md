# 체이닝 API 문법 전수조사: V-기준 대비 검증 결과

작성일: 2026-07-10
근거 문서: `../reference/query_api_language_survey_202607.md` (검증 기준 V1~V9의 출처)
상태: **리팩토링 완료 (2026-07-10 전면 승인 - 7장 실행 결과 참조)**
검증 방법: 코드 정독 판독. 실행으로 재현하지 않은 판정은 "(실행 미검증)" 표기 -
실행 고정은 후속 핀 테스트(C0)의 몫.

## 1. 전수 대상 (grep 검산 완료)

`^Public (Function|Sub|Property)` 전수 대조로 누락 0 확인 (2026-07-10, HEAD c643092):

| 표면 | Public 멤버 수 | 검증 대상 | 제외 (사유) |
|---|---|---|---|
| UnivDB.bas (파사드) | 11 | 11 | - |
| clsQuery.cls | 12 | 12 | - |
| clsReportBuilder.cls | 15 | 15 | - |
| clsTableBuilder.cls | 8 | 8 | - |
| clsTable.cls | 33 | 13 (명령 표면) | 20: 인프라/조회 표면 (Load*/Get*/SelectRows*/GetUniqueKeys*/Generation/IndexBackend/RowCount/ColCount/Headers/RawArray/Clear/GetValue/GetRecord) - shim 동결 + 명령어 문법 스코프 밖 |

검증 대상 합계 59개 멤버.

## 2. 검증 매트릭스

셀 표기: `O` pass / `X` violation (F-번호 참조) / `-` n/a (기준이 해당 없음).

### 2.1 UnivDB.bas (파사드)

| 멤버 | V1 | V2 | V3 | V4 | V5 | V6 | V7 | V8 | V9 |
|---|---|---|---|---|---|---|---|---|---|
| InitDB (Sub) | - | - | O | O | - | - | O | - | O |
| RegisterTable (Sub) | - | O | O | O | - | - | O | - | O |
| GetTable | - | O | O | O | O | - | O | - | O |
| CreateTable | - | - | O | O | O | - | O | - | O |
| Query | - | - | O | X F15 | O | - | O | - | O |
| Report | - | - | O | X F15 | O | - | O | - | O |
| Join | - | O | O | O | X F17 | - | O | - | X F21 |
| Upsert (Sub) | - | O | O | O | - | - | O | - | O |
| Builder | - | - | O | X F15 | O | - | O | - | O |
| EditTable | - | O | O | X F15 | O | - | O | - | O |
| ClearDB (Sub) | - | - | O | O | - | - | O | - | O |

### 2.2 clsQuery.cls

| 멤버 | V1 | V2 | V3 | V4 | V5 | V6 | V7 | V8 | V9 |
|---|---|---|---|---|---|---|---|---|---|
| From | - | O | O | O | O | - | O | - | O |
| Where | - | O | O | X F11주1 | O | - | O | - | O |
| Filter | X F1, F6 | X F5 | X F7 | O | O | X F18 | O | - | X F21 |
| SelectColumns | - | O | X F9 | O | O | - | O | X F20 | O |
| OrderBy | - | O | X F8 | O | O | - | O | - | O |
| Top | - | O | O | O | O | - | O | - | O |
| SetValue | - | O | O | X F11 | O | - | O | - | O |
| Update | - | O | O | O | O | - | O | - | O |
| Delete | - | O | O | X F12 | O | - | O | - | O |
| Insert | - | O | O | O | O | - | O | - | O |
| SelectTable | - | - | O | X F10 | O | - | O | - | O |
| DebugPrint (Sub) | - | - | O | O | - | - | O | - | O |

주1: Where 자체는 동등 전용이라 연산자 인자가 없어 V1 무관. V4의 F11은 SetValue와의
어휘 계열 문제로, Where는 인덱스 고속 경로임이 문서화돼 있어 pass로 판정.

### 2.3 clsReportBuilder.cls

| 멤버 | V1 | V2 | V3 | V4 | V5 | V6 | V7 | V8 | V9 |
|---|---|---|---|---|---|---|---|---|---|
| Init | - | O | O | O | O | - | O | - | O |
| AddColumns | - | O | O | X F13 | O | - | O | X F20 | O |
| RemoveColumns | - | O | O | X F13 | O | - | O | X F20 | O |
| CreateKey | - | O | O | O | O | - | O | - | O |
| JoinTable | - | O | O | O | O | - | O | - | O |
| FillFromMap | - | O | O | O | O | - | O | - | O |
| ConvertDates | - | O | O | O | O | - | O | X F20 | O |
| Filter | X F1(상속) | X F5 | X F7 | O | O | X F18 | O | - | X F21 |
| Exclude | X F1(상속) | X F5 | O | X F12 | O | X F18 | O | - | O |
| UpdateValue | X F1(상속) | X F5 | O | X F11 | O | X F18 | O | - | O |
| Table (Get) | - | - | O | X F10 | O | - | O | - | O |
| SelectColumns | - | O | X F9 | O | O | - | O | X F20 | O |
| SelectFinal | - | O | O | X F14 | O | - | O | X F20 | O |
| DefineColumn | - | O | O | O | O | - | O | - | O |
| BuildToSheet (Sub) | - | O | O | O | O | - | O | - | O |

### 2.4 clsTableBuilder.cls

| 멤버 | V1 | V2 | V3 | V4 | V5 | V6 | V7 | V8 | V9 |
|---|---|---|---|---|---|---|---|---|---|
| KeyColumn (Let) | - | O | O | O | - | - | O | - | O |
| InitFromTable (Sub) | - | O | O | O | - | - | O | - | O |
| Row | - | O | O | O | O | - | O | - | O |
| ExtractRow | - | O | O | O | O | - | O | - | O |
| Val | - | X F5주2 | O | X F11 | X F19 | - | X F19 | - | O |
| PatchIf | X F4 | X F5 | O | O | O | X F18 | O | - | O |
| ToTable | - | - | O | O (canonical) | O | - | O | - | O |
| Clear (Sub) | - | O | O | O | - | - | O | - | O |

주2: Val getter는 미지 컬럼에 Empty 반환(148~164) - 컬럼 오타가 조용히 Empty로
읽히는 관용 처리. setter는 자동 스키마 확장(169~171)이라 의도된 동작이나 getter의
조용한 Empty는 오타 검출 불가.

### 2.5 clsTable.cls (명령 표면 13)

| 멤버 | V1 | V2 | V3 | V4 | V5 | V6 | V7 | V8 | V9 |
|---|---|---|---|---|---|---|---|---|---|
| GroupBy | X F2 | X F5 | O | O | O | O | O | X F20 | X F22 |
| OrderBy (Sub) | - | O | X F8 | O | X F16 | - | O | - | O |
| Compute (Sub) | X F3 | X F3, F5 | O | O | X F16 | X F18 | O | - | X F22 |
| SelectColumns | - | O | X F9 | O | O | - | O | X F20 | O |
| UpsertFrom (Sub) | - | O | O | O | X F16 | - | O | - | O |
| UpdateBatch (Sub) | - | O | O | X F11계열 | X F16 | - | O | - | O |
| DeleteBatch (Sub) | - | O | O | X F12 | X F16 | - | O | - | O |
| AddColumn (Sub) | - | O | O | X F13 | X F16 | - | O | X F20 | O |
| RemoveColumn (Sub) | - | O | O | X F13 | X F16 | - | O | X F20 | O |
| CopyColumn (Sub) | - | O | O | O | X F16 | - | O | - | O |
| DumpToSheet (Sub) | - | O | O | O | - | - | O | - | O |
| Value (Let) | - | O | O | X F11 | - | - | O | - | O |
| Value (Get) | - | O | O | O | - | - | O | - | O |

## 3. 위반 목록 (F-1 ~ F-22, 증거 포함)

### 등급 S1 - 조용한 실패 (fail-fast 부재, 최우선)

- **F1 [V1] clsQuery.Filter 미지 연산자 = 조용한 0행**: 바깥 Select Case
  (clsQuery.cls:130~206)에 Case Else 없음. 오타(`"=>"` 등)면 어느 분기도 안 타고
  빈 setSurvivors가 그대로 후보에 대입(208) - 에러 없이 0행. (실행 미검증)
  clsReportBuilder.Filter(148)/Exclude(156)/UpdateValue(164)가 sOp를 그대로 위임해
  동일 결함 상속.
- **F2 [V1] clsTable.GroupBy 미지 집계함수 = 조용한 빈 컬럼**: 누적 분기
  (clsTable.cls:639 단일 경로, 741 복합 경로)에 Case Else 없음 - 미지 함수(`"SUMM"`
  오타)는 아무것도 누적하지 않고, 렌더 fallback `Case Else ' SUM, MAX, MIN`(680, 813)이
  빈 vStats를 그대로 기입. 결과: 헤더(`Col_SUMM`)는 생성되고 값은 전부 Empty.
  (실행 미검증. 사전 조사의 "SUM으로 오계산" 추정은 코드 재판독으로 정정 -
  정확한 양식은 "조용한 빈 컬럼")
- **F3 [V1][V2] clsTable.Compute 3중 조용한 실패**: (1) 미지 연산자 - Select Case
  (1283~1289) Case Else 없음, 조용한 무동작. (2) 피연산자 컬럼 오타 - 조용한
  `Exit Sub`(1271, 1274). 이때 타겟 컬럼은 이미 자동 생성된 채(1268) 빈 컬럼으로
  잔류 - 구조 변경은 하고 값은 안 채우는 반쪽 실행. (3) `On Error Resume Next`(1282)가
  행별 타입 에러를 삼킴.
- **F4 [V1] clsTableBuilder.PatchIf 미지 연산자 = 조용한 무매칭**: `Case Else:
  bMatched = False`(clsTableBuilder.cls:193) + `On Error Resume Next`(185)가 비교
  에러도 무매칭으로 흡수.
- **F5 [V2] 오류 처리 3단 비대칭**: 컬럼 오타는 Err.Raise 9999(clsQuery.cls:123,
  GroupBy도 580/그룹 컬럼 동일) / 연산자 오타는 조용한 0행·무동작(F1~F4) / Compute는
  컬럼 오타조차 조용(F3). 같은 등급의 입력 오류가 표면마다 다른 방식으로 실패.
  clsTableBuilder.Val getter의 미지 컬럼 Empty 반환(162~163)도 이 계열.
- **F6 [V1] clsQuery.Filter 수치 경로 비숫자 비교값 = 조용한 0행**: clsQuery.cls:139~144.
  `.Filter("Age", ">", "삼십")`이 에러 없이 빈 결과. (실행 미검증)

### 등급 S2 - 동명이의 (같은 이름, 다른 의미론)

- **F7 [V3] Filter**: clsQuery.Filter(106)는 지연 - rowID 후보 축소만, 물질화 없음.
  clsReportBuilder.Filter(148)는 즉시 - 내부에서 `UnivDB.Query(m_tWork)...SelectTable()`을
  실행해 작업 테이블을 교체. 같은 시그니처 `(col, op, val)`, 다른 실행 시점.
- **F8 [V3] OrderBy**: clsQuery.OrderBy(283)는 Function(clsQuery 반환, 지연 정렬 스펙).
  clsTable.OrderBy(831)는 Sub(즉시 in-place 정렬 + Generation 증가). 같은 이름이
  체인 가능 여부까지 다름.
- **F9 [V3] SelectColumns**: clsQuery(219, 지연 투영 스펙) vs clsReportBuilder(218,
  즉시 작업본 축소, Me 반환) vs clsTable(1178, 즉시 새 테이블 반환). 3클래스 3의미.

### 등급 S3 - 어휘·규약 비일관

- **F10 [V4] 종결자 어휘 3종**: SelectTable(clsQuery:427) / Table(Property Get,
  clsReportBuilder:209) / ToTable(clsTableBuilder:204). survey 축 D 근거로 canonical
  후보는 ToTable(LINQ To* 규약 정합).
- **F11 [V4] 값 설정 어휘 4종**: SetValue(clsQuery:302) / UpdateValue(clsReportBuilder:164)
  / Val(clsTableBuilder:148) / Value Let(clsTable:1064). UpdateBatch(clsTable:978)도 계열.
- **F12 [V4] 행 제거 어휘 3종 + 반환 3종**: Delete(Long, clsQuery:339) / Exclude(Me,
  clsReportBuilder:156) / DeleteBatch(Sub, clsTable:1006).
- **F13 [V4] 컬럼 추가·삭제 단복수 혼재**: AddColumns/RemoveColumns(csv 복수,
  clsReportBuilder:57/67) vs AddColumn/RemoveColumn(단건, clsTable:1104/1133).
- **F14 [V4] SelectFinal 레거시 별칭**: clsReportBuilder:233이 SelectColumns(218)로
  위임하는 순수 별칭 - 같은 클래스 안 같은 기능 두 이름, deprecated 표기 없음.
- **F15 [V4] 진입점 품사 혼재**: Query/Report/Builder(명사) vs EditTable(동사구,
  UnivDB.bas:131). 경미.
- **F18 [V6] 연산자 정규화 3종**: `UCase(Trim(sOp))`(clsQuery:130) / `Trim(Operator)`
  대소문자 유지(clsTableBuilder:186) / raw `sOp`(clsTable:1283). 집계함수는
  `UCase(Trim())`(clsTable:571). 같은 오타가 표면마다 다르게 실패하는 원인.
- **F20 [V8] 목록 인자 규약 혼재**: csv 문자열(AddColumns/RemoveColumns/GroupBy의
  sGroupCols·sAggList) vs ParamArray(SelectColumns/ConvertDates/SelectFinal).
  clsReportBuilder 한 클래스 안에서도 혼재(AddColumns csv vs SelectColumns ParamArray).

### 등급 S4 - VBA 제약 누출·체인 단절

- **F16 [V5] clsTable 변형 메서드 8개가 Sub**: OrderBy(831)/Compute(1262)/UpsertFrom(866)
  /UpdateBatch(978)/DeleteBatch(1006)/AddColumn(1104)/RemoveColumn(1133)/CopyColumn(1081).
  GroupBy/SelectColumns(Function, clsTable 반환)와 비일관 - `t.GroupBy(...).OrderBy(...)`
  체인이 컴파일 불가.
- **F19 [V7] clsTableBuilder.Val**: getter(Variant)/setter(`Set Val = Me`, 174) 혼성
  Function - 문(statement) 위치 괄호 호출이 구문 오류라 **Call 강제**
  (TestTableBuilder.bas:56, 59에 실증). VBA 내장 `Val()` 섀도잉까지 겹침.
- **F17 [V5] 파사드 Join 체인 단절 (관찰)**: Join이 clsTable 반환(UnivDB.bas:93)이라
  조인 후 재쿼리는 `UnivDB.Query(joined)` 재래핑 필요. 위반이라기보다 인체공학 관찰 -
  clsQuery에 Join 부재가 원인.

### 등급 S5 - 계약 문서 공백

- **F21 [V9] OR 부재 미문서화**: 조건 결합이 AND 전용(Where=Intersect, Filter=후보
  축소)임이 UnivDB_Reference.md에 미기재. FilterIn이 동일 컬럼 OR을 커버한다는 안내도
  없음. 교차 컬럼 OR은 표현 불가.
- **F22 [V9] 암묵 동작 미문서화**: Compute 상수 피연산자 미지원(코드 주석
  clsTable.cls:1273에만 존재), GroupBy 집계함수 생략 시 COUNT 기본(574) - 공개 문서
  미기재.

## 4. 기준별 집계

| 기준 | 위반 F | 위반 멤버 수 (상속 포함) |
|---|---|---|
| V1 미지 문자열 즉시 raise | F1 F2 F3 F4 F6 | 7 (Filter, RB Filter/Exclude/UpdateValue, GroupBy, Compute, PatchIf) |
| V2 오류 처리 대칭성 | F5 (F3, F6 연동) | 상동 + Val getter |
| V3 같은 이름 같은 의미론 | F7 F8 F9 | 3쌍 7멤버 |
| V4 같은 개념 같은 이름 | F10~F15 | 15+ |
| V5 반환 타입 예측 가능성 | F16 (F17 관찰, F19 연동) | 9 |
| V6 정규화 단일 지점 | F18 | 4 표면 |
| V7 VBA 제약 비노출 | F19 | 1 (Val) |
| V8 인자 규약 일관성 | F20 | 10+ |
| V9 표현력 공백 문서화 | F21 F22 | 문서 2건 |

심각도 순위 (프로젝트 철학 기준):
**S1 조용한 실패(F1~F6) > S2 동명이의(F7~F9) > S3 어휘 비일관(F10~F15, F18, F20)
> S4 체인 단절·VBA 누출(F16, F19) > S5 문서 공백(F21, F22)**

## 5. 다음 단계 골격 (게이트: 개선 착수 범위는 사용자 결정)

개선안 상세 도출은 이 문서의 스코프 밖. 위반-개선 후보 매핑의 골격만 남긴다:

- **S1 → fail-fast화** (계약 변경, D0~D3 검증된 2단 커밋 패턴: 핀 고정 후 flip):
  단일 연산자 검증 헬퍼 신설(F18 선행 해소) → F2(GroupBy) → F1 → F3 → F4 → F6 순.
  F6은 "raise vs 정당한 no-match" 계약 결정 필요.
- **S2/S3 → additive 어휘 통일** (기존 이름 영구 유지 + canonical 별칭): ToTable
  수렴(F10), Keep 별칭으로 즉시/지연 어휘 분리(F7), GetVal/SetVal 분리(F19, F11),
  deprecated 표기(F14).
- **S4 → 체인화**: clsTable 변형 Sub의 Function(Me 반환) 전환 여부(F16) - 시그니처
  변경 취급이라 별도 승인 필요.
- **S5 → Reference 갱신**: AND 전용·암묵 기본값·상수 미지원 명문화(F21, F22).

## 6. 스코프 밖 (명시)

- clsTable 인프라/조회 표면 20개 멤버 (shim 동결 - 금지 사항).
- 교차 컬럼 OR 도입, Compute 상수 피연산자 등 표현력 확장 - 위반 아님(공백),
  별도 task로 이월 판단.
- jOOQ류 코드 생성 타입 안전화 - VBA 도구 체인 부재 (survey 4장).

## 7. 실행 결과 (2026-07-10 전면 리팩토링 완료)

"전수조사 기반 전면 리팩토링" 승인으로 S1~S5 전 등급 실행. 2단 커밋 패턴
(C0 핀 고정 -> 계약 변경 + 핀 flip 동일 커밋), 매 커밋 vbaflow test 전량 green.

### 7.1 추보 발견 (매트릭스 정정 2건)

C3 구현 중 코드 재정독으로 전수조사 매트릭스의 오판정 2건을 발견, 정정한다:

- **F23 [V1/V2] clsTable.CopyColumn 소스 컬럼 오타 = 조용한 무동작** (구 L1089
  `Exit Sub`). 매트릭스 2.5의 CopyColumn V2=O는 오판정.
- **F24 [V2] clsTable.Value(Let) 미지 컬럼 = Debug.Print 경고 후 무시** (구 L1080) -
  Debug.Print는 헤드리스에서 안 보이므로 사실상 조용한 무시. 매트릭스 2.5의
  Value(Let) V2=O는 오판정.

두 건 모두 fail-fast(9999)로 교정 완료. 교훈: "코드 판독 (실행 미검증)" 매트릭스는
구현 단계의 재정독을 검산으로 삼는다.

### 7.2 커밋 이력

- `da44338` **C0 핀**: TestChainingSemantics.bas 신설, F1~F6 현행 조용한 실패 고정.
- `37375ef` **C1a**: modOperators 신설(화이트리스트 단일 정규화 지점, F18 해소) +
  clsQuery.Filter 미지 연산자/비숫자 비교값 raise (F1, F6). RB 3메서드 자동 상속.
- `20a209b` **C1b**: GroupBy 미지 집계함수 raise (F2) - 파싱 단일 지점이라 단일/복합
  경로 동시 커버. 함수 생략 시 COUNT 기본은 유효 계약 유지.
- `ab9e9b6` **C1c**: Compute 문장 수준 검증을 타겟 컬럼 생성 앞으로 재배치 (F3) -
  반쪽 실행 제거. 행 데이터 수준 오류는 계약된 관용 유지 (문장/데이터 등급 분리).
- `bb1e056` **C1d**: PatchIf 미지 연산자 raise (F4).
- `e093aee` **C2**: 어휘 통일 (additive) - ToTable canonical 종결자 (F10, clsQuery/RB),
  Keep 즉시 실행 어휘 (F7), SetVal/GetVal 분리 (F19, GetVal은 fail-fast getter),
  SelectFinal deprecated 표기 (F14). 기존 이름 전부 영구 유지.
- `0f7ec38` **C3**: clsTable 변형 8메서드 Sub -> Function(Me 반환) 전환 (F16) - 문
  위치 호출 소스 호환, 기존 스위트 전량 green으로 무파손 확인. + F23/F24 교정.
- **C4 (본 커밋)**: UnivDB_Reference.md v2.7.0 - 6.5 연산자 화이트리스트/fail-fast 계약,
  6.6 표현력 공백 명문화 (F21, F22), canonical 어휘/실행 시점 주의/멱등 계약/VBA 체인
  구문 계약 반영. 본 문서 상태 갱신.

### 7.3 확정된 계약 요약

- **문장 수준 오류는 raise 9999** (미지 연산자/집계함수, 컬럼명 오타, 수치 연산자에
  비숫자 비교값), **데이터 수준 문제는 NULL/skip 계약** (Empty/Error/비숫자 셀 불참).
- **canonical 어휘**: ToTable(종결) / Keep·Exclude(즉시 실행 필터) / SetVal·GetVal
  (빌더 값 접근). 레거시 이름(SelectTable/Table/Filter(RB)/SelectFinal/Val)은 영구 유지.
- **clsTable 변형 메서드는 전부 체인 가능** (Me 반환, in-place 의미론/세대 계약 불변).
- **멱등 계약 명문화**: AddColumn=ensure, RemoveColumn=drop, DeleteBatch 0건=무동작.

### 7.4 잔여 이월 (미착수 확정)

- F8/F9 (OrderBy/SelectColumns 동명이의): 의미 차이는 Reference에 문서화로 완화.
  이름 분리는 이득 대비 표면 증식이 커서 보류.
- ~~F11~F13, F15, F20 (잔여 어휘 통일)~~: **2026-07-10 B축으로 해소** - 7.6절 참조.
- C-1 (교차 컬럼 OR), C-2 (Compute 상수): 표현력 확장 축 - 별도 task로 이월.
- ~~F17 (파사드 Join 체인 단절)~~: **2026-07-10 해소** - 7.5절 참조.

### 7.5 F17 해소 + 추보 F25 (2026-07-10)

- `1ce5f27` **clsQuery.Join 신설** (F17): `Join(vRight, sLeftKey, [jType=jtInner],
  [sRightPrefix], [sRightKey]) As clsQuery`. 의미론은 **물질화 장벽** - Join 시점까지
  스테이징된 조건/투영/정렬/Top을 ToTable()로 실행해 좌측 확정 후 조인 결과로 재소싱
  (From + ResetPipelineState). 지연 엔진이 단일 소스 전제라 지연 조인은 불가 - 장벽
  의미론을 @desc/Reference에 명시해 V3(같은 표면 내 실행 시점 혼재)을 계약으로 완화.
  기본 jType은 파사드 정합(jtInner), ReportBuilder.JoinTable(jtLeft)과의 차이 문서화.
- `a128fa3` **추보 F25 [V1/V2 + 헤드리스]**: `modEngine.JoinTable` 키 컬럼 부재 처리가
  MsgBox + Exit Function(Nothing 반환 -> 조용한 91 전이)이었다 - 헤드리스에서 모달 행
  유발이라 **현행 핀 고정이 불가능**해 2단 커밋 패턴 예외(fix + 신규 계약 핀만).
  Err.Raise 9999로 교정. 전수조사 매트릭스가 modEngine(비체이닝 엔진 모듈)을 스코프
  밖에 둬 놓친 사례 - 엔진 모듈의 오류 처리도 후속 점검 후보.

### 7.6 B축 잔여 어휘 통일 실행 결과 (2026-07-10)

원칙: 개념이 진짜 같은 곳만 canonical 별칭 수렴(additive), 개념이 다른 것의 이름
통일은 V3 역행이라 어휘 지도 문서화. 별칭 범위는 사용자 확정 3종.

- `6581ae7` **F11 부분 해소**: clsQuery.SetVal(=SetValue - 빌더 SetVal과 같은 개념),
  RB.PatchIf(=UpdateValue, 빌더 PatchIf 인자 순서 cond/op/val/tgt/new로 정합).
  clsTable.Value(Let)는 직접 셀 쓰기로 개념 상이 - 무접촉.
- `058d3a8` **F13/F20 해소**: SelectColumns의 플랫 파싱을 FlattenColumnArgs로 추출
  공용화, clsTable.AddColumns/RemoveColumns 신설(개별/csv/Array 허용, 멱등 위임).
  목록 인자 canonical 규약 = "ParamArray + csv/Array 허용", RB csv 전용은 레거시.
- **문서화만으로 확정 (재론 시 이 결정 인용)**: F12 행 제거 3종은 서로 다른 연산
  (스테이지드 CRUD/술어 즉시/동등 일괄)이라 별칭 없음. F15 UnivDB.Editor 별칭은
  미채택(경미). Reference v2.7.2에 "6.7 어휘 지도(Vocabulary Map)" 신설로 전 개념 x
  표면 대응을 명문화.

이로써 audit 위반 F1~F25 중 코드 접촉 대상은 전부 종결. 잔여는 표현력 확장 축
(C-1 교차 컬럼 OR, C-2 Compute 상수)과 F8/F9 의도적 보류(문서화 완화)뿐.
