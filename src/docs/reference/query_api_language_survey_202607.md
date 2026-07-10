# 상용 쿼리 API 언어철학 조사와 UnivDB 문법 검증 기준

작성일: 2026-07-10
용도: 체이닝 API 문법 검증(`../tasks/2026-07-10_chaining_api_audit.md`)의 근거 자료
검증 방법: 에이전트 지식으로 초안 작성 후 핵심 주장을 웹 검색으로 검증.
출처가 붙은 주장은 검증됨, "(에이전트 지식, 미검증)" 표기는 검증 생략분.

## 0. 목적과 재수행 경위

과거 대화에서 동일 주제 조사가 수행됐으나 문서로 박제되지 않아 유실됐다. 이번 조사는
처음부터 파일로 남긴다. 목적은 두 가지다:

1. 상용 DB/쿼리 API 6종의 언어 설계 철학에서 **문법 검증 기준(V-기준)** 을 도출한다.
2. 그 기준으로 UnivDB의 메서드 체이닝 명령어 표면을 전수조사한다 (별도 audit 문서).

UnivDB의 판단 기준("조용히 틀린 결과가 최악의 실패 양식", fail-fast 우선, 배포용
라이브러리 지향)과 정합하는 기준만 채택하고, 각 기준에는 어느 제품의 어느 규율에서
왔는지 근거를 남긴다.

## 1. 비교 대상 6종 프로필

### 1.1 SQL (선언형 기준점)

- 무엇을: 선언형 쿼리 언어의 원점. 문법과 카탈로그(스키마)가 실행 전 검증의 계약.
- 어떻게: 미지 식별자(컬럼/테이블)나 미지 함수는 **항상 파스/바인드 단계 에러**다.
  `SELECT SUMM(x)`는 절대 SUM으로 해석되지 않는다. (에이전트 지식, 미검증 - 모든
  주류 RDBMS 공통 동작이라 검증 생략)
- 왜: 쿼리가 데이터를 만지기 전에 문장 자체의 유효성이 판정된다. "실행됐다 = 문법과
  이름은 전부 유효했다"는 계약이 사용자 신뢰의 바닥이다.
- 반례 메모: MySQL의 관용 모드(느슨한 GROUP BY, zero-date 허용 등)는 strict mode
  도입으로 후퇴한 역사가 있다 - 관용은 언젠가 strict로 회귀한다. (에이전트 지식, 미검증)

### 1.2 LINQ (C# 메서드 체이닝 - UnivDB와 가장 가까운 사촌)

- 무엇을: 컬렉션 위 메서드 체이닝 쿼리. UnivDB clsQuery와 동형의 구조.
- 어떻게: 표준 연산자를 **실행 방식으로 공식 분류**한다 - 지연(deferred: Where/Select,
  streaming/non-streaming 세분) vs 즉시(immediate: Count/Max 등 스칼라 반환 전부).
  `ToList`/`ToArray` 같은 **`To*` 접두 메서드가 "지금 물질화한다"는 종결 신호**다.
  출처: [Microsoft Learn - Classification of Standard Query Operators by Manner of
  Execution](https://learn.microsoft.com/en-us/dotnet/visual-basic/programming-guide/concepts/linq/classification-of-standard-query-operators-by-manner-of-execution)
- 왜: 사용자가 메서드 이름만 보고 "이 호출이 소스를 읽는가"를 예측할 수 있게 한다.
  실행 시점이 이름 규약에 인코딩되어 있다.

### 1.3 dplyr / tidyverse (동사 문법과 이름 규율)

- 무엇을: "데이터 조작의 문법". 소수의 동사(mutate/select/filter/summarize/arrange)를
  파이프로 연결한다.
- 어떻게: 함수명은 **명령형 동사**(mutate, not mutated/mutating), 같은 접두사로 관련
  기능을 묶고, 인자 순서(데이터가 항상 첫 인자)까지 전 패키지 공통 규약. "하나를
  배우면 다음 것이 쉬워진다"가 명시 목표.
  출처: [Tidy design principles - Function names](https://design.tidyverse.org/function-names.html),
  [The tidy tools manifesto](https://tidyverse.tidyverse.org/articles/manifesto.html)
- 왜: 인간 중심 설계 - API의 일관성이 곧 학습 비용과 오용률이다. 같은 개념에 같은
  이름, 같은 이름에 같은 의미론이 규율의 핵심이다.

### 1.4 jOOQ (타입 안전 fail-fast의 대표)

- 무엇을: Java용 SQL DSL. 스키마에서 클래스를 생성해 쿼리를 타입 안전하게 조립.
- 어떻게: 연산자와 컬럼이 문자열이 아니라 **생성된 메서드**(`field.eq(...)`,
  `field.gt(...)`)다. 컬럼명 오타·타입 불일치가 **컴파일 타임에 잡힌다** - 문자열 SQL은
  "문장마다 테스트를 써야 유효성이 보장"되는 반면 jOOQ는 그 간극을 컴파일러가 메운다.
  출처: [Type Safe SQL in Java with jOOQ](https://martinelli.ch/type-safe-sql-in-java/),
  [jOOQ blog - why execute jOOQ queries with jOOQ](https://blog.jooq.org/why-you-should-execute-jooq-queries-with-jooq/)
- 왜: 오류 검출 시점을 최대한 앞으로 당기는 것이 설계의 존재 이유. 비용은 코드 생성
  파이프라인 - VBA엔 이 경로가 없으므로 "검출 시점 앞당기기"라는 원칙만 가져온다.

### 1.5 Knex / Kysely (문자열 연산자 - UnivDB 현행과 동형)

- 무엇을: JS/TS 쿼리 빌더. `where('col', '>', 5)`처럼 **연산자를 문자열 인자로** 받는다 -
  UnivDB `Filter(col, op, val)`과 정확히 같은 형태.
- 어떻게: 문자열 연산자를 **화이트리스트로 검증하고, 미지 연산자는 즉시
  `The operator "X" is not permitted` 예외**를 던진다 (knex 내부 wrappingFormatter의
  허용 목록). 조용히 0행을 반환하는 경로가 없다.
  출처(간접 실증): [mikro-orm issue #3647 - TypeError: The operator "is" is not
  permitted](https://github.com/mikro-orm/mikro-orm/issues/3647) (knex 계층에서 발생)
- OR 표현: `orWhere(...)` 메서드와 그룹핑 콜백(`where(function() {...})`)으로 AND/OR
  결합을 제공한다. (에이전트 지식, 미검증)
- 왜: 문자열 연산자 방식 자체는 동적 언어에서 합리적 선택이다. 단 그 대가로 **미지
  문자열 즉시 거부**가 필수 규율이 된다 - 문자열의 자유도와 fail-fast는 한 쌍이다.

### 1.6 pandas (반면교사 - 관용의 실패 양식)

- 무엇을: 즉시 실행 데이터프레임 API. 표현력은 크지만 오류 철학이 관용 쪽이다.
- 어떻게: NaN이 연산을 타고 **조용히 전파**되고, `to_numeric(errors="coerce")`는 변환
  실패를 NaN으로 삼키며 특정 조건에선 coerce 자체가 조용히 실패한 사례들이 이슈로
  보고돼 있다. 실무 가이드들이 "명시적 dtype 지정 + 경계 검증"을 방어책으로 권한다.
  출처: [pandas user guide - Working with missing data](https://pandas.pydata.org/docs/user_guide/missing_data.html),
  [pandas issue #32394](https://github.com/pandas-dev/pandas/issues/32394),
  [pandas issue #24910](https://github.com/pandas-dev/pandas/issues/24910)
- 왜(반면교사): "대부분의 데이터 버그는 크래시하지 않고 속삭인다." 관용 모드의 비용은
  틀린 숫자가 보고서까지 도달하는 것이다 - UnivDB 철학이 막으려는 바로 그 양식.

## 2. 분석 축과 UnivDB 시사점

### 축 A. 지연 vs 즉시 실행, 그리고 그 전달 방식

LINQ는 실행 시점을 공식 분류하고 이름 규약(`To*`, 스칼라 반환=즉시)으로 전달한다.
Spark의 transformation/action 이분법도 같은 원리다 (에이전트 지식, 미검증).
같은 이름의 메서드가 문맥에 따라 지연/즉시로 갈리는 설계는 어느 제품에도 없다.

**UnivDB 시사점 [SV-A]**: 실행 시점은 API 표면에서 예측 가능해야 한다. 같은 이름
`Filter`가 clsQuery에선 지연(후보 축소), clsReportBuilder에선 즉시(머티리얼라이즈)인
것은 이 축의 정면 위반. 최소한 계약 문서에 실행 시점을 명기하고, 장기적으로 어휘를
분리해야 한다.

### 축 B. 연산자 표현: 문자열 vs 메서드 vs 표현식 객체

- 메서드형(jOOQ `eq/gt`): 오타가 컴파일 에러. VBA에선 컬럼이 동적(런타임 문자열)이라
  코드 생성 없이는 불가능하고, 메서드 오버로드도 없어 조합 폭발이 난다.
- 표현식 객체: 람다/연산자 오버로드 필요 - VBA에 없음.
- 문자열형(Knex/Kysely, UnivDB 현행): VBA 제약 하에서 **합리적 선택**. 단 Knex가
  보여주듯 필수 규율이 따라온다: **(1) 미지 문자열 즉시 raise, (2) 정규화(대소문자,
  Trim)는 단일 지점에서 한 번.**

**UnivDB 시사점 [SV-B]**: 문자열 연산자 유지는 정당하다. 결함은 문자열이라는 선택이
아니라 화이트리스트 거부의 부재다. 연산자/집계함수 문자열을 받는 모든 표면은 단일
정규화 헬퍼를 거쳐 미지 값에서 Err.Raise해야 한다.

### 축 C. 오류 철학: fail-fast vs 관용

jOOQ/SQL strict는 오류 검출을 앞으로 당기고, pandas류 관용은 뒤로(때로는 영원히)
미룬다. 관용의 실패 양식은 "조용히 틀린 숫자" - UnivDB가 명시적으로 최악으로 규정한
것. 그리고 오류 처리의 **비대칭**(어떤 오타는 raise, 어떤 오타는 무시)은 사용자의
신뢰 모델을 깬다 - "컬럼 오타는 잡아주니 연산자 오타도 잡아주겠지"라는 추론이
배신당한다.

**UnivDB 시사점 [SV-C]**: 입력 오류의 처리 등급은 표면 전체에서 대칭이어야 한다.
현행은 컬럼명 오타 raise(9999) vs 연산자 오타 조용한 0행 vs 집계함수 오타 조용한
오계산으로 3단 비대칭 - 가장 시급한 교정 대상.

### 축 D. 종결자(터미널) 명명 규약

LINQ `To*`(ToList/ToArray/ToDictionary), dbplyr `collect()`, Spark action처럼 성숙한
API는 "체인을 끝내고 물질화한다"를 이름으로 신호한다. UnivDB의 세 빌더는 같은 개념을
`SelectTable()` / `Table`(Property) / `ToTable()` 세 이름으로 부른다.

**UnivDB 시사점 [SV-D]**: 종결자 어휘는 하나로 수렴해야 하며, 업계 규약(`To*` 접두)과
이미 정합인 `ToTable`이 canonical 후보다. 기존 이름은 호환 별칭으로 유지(shim 시그니처
변경 금지 제약과도 정합).

### 축 E. VBA 제약 매핑 (제약이 배제/강제하는 설계)

| VBA 제약 | 배제되는 설계 | 강제되는 규율 |
|---|---|---|
| 람다/연산자 오버로드 없음 | 표현식 객체 DSL | 문자열 연산자 + 화이트리스트 (축 B) |
| 메서드 오버로드 없음 | eq/gt 메서드형의 타입별 변형 | Variant 인자 + 런타임 검증 |
| Sub는 체인 불가 | - | 변형 메서드는 Function(Me/신객체 반환)이어야 체인 유지 |
| 값 반환 Function을 문(statement) 위치에서 괄호 호출하면 구문 오류 | getter/setter 혼성 Function 체이닝 | `Call` 강제를 유발하는 혼성 반환 금지, 역할 분리 |
| 내장 함수 이름 섀도잉 가능 (`Val` 등) | - | 내장과 충돌하는 멤버명 회피 |
| ParamArray는 명명 인자 불가 | - | 인자 규약(csv vs ParamArray)을 클래스 간 통일 |

**UnivDB 시사점 [SV-E]**: `clsTableBuilder.Val`의 Call 강제는 VBA 제약을 표면에 누출한
사례(혼성 반환 Function + 괄호 체이닝). clsTable.OrderBy/Compute가 Sub인 것은 체인
단절의 직접 원인. 제약 표의 오른쪽 열이 곧 검증 항목이다.

## 3. 결론: 문법 검증 기준 (V-기준)

아래 9개 기준으로 체이닝 명령어 표면을 전수 검증한다. 각 기준의 근거는 위 시사점 ID.

| ID | 기준 | 근거 |
|---|---|---|
| V1 | 미지 연산자/함수 문자열은 즉시 raise - 조용한 통과(0행/무동작/오계산) 금지 | SV-B, SV-C (Knex 화이트리스트, SQL 원점 계약, pandas 반면교사) |
| V2 | 오류 처리 대칭성 - 같은 등급의 입력 오류(이름 오타류)는 같은 방식으로 실패 | SV-C |
| V3 | 같은 이름은 같은 의미론 - 실행 시점·부수효과 포함 (동명이의 금지) | SV-A (LINQ 실행 분류), dplyr 규율 |
| V4 | 같은 개념은 같은 이름 - 종결자/값 설정/삭제 어휘의 클래스 간 통일 | SV-D, dplyr "하나 배우면 다음이 쉽다" |
| V5 | 체인 반환 타입 예측 가능성 - 변형은 체인 유지, 종결은 이름으로 신호(To* 류) | SV-D, SV-E (Sub 체인 불가) |
| V6 | 문자열 인자 정규화의 단일 지점 - Trim/UCase 규약이 표면마다 다르면 안 됨 | SV-B |
| V7 | VBA 제약의 표면 비노출 - Call 강제, getter/setter 혼성 반환, 내장 이름 충돌 금지 | SV-E |
| V8 | 인자 전달 규약 일관성 - 같은 종류의 목록 인자는 같은 방식(csv vs ParamArray 혼재 금지) | SV-E, dplyr 인자 규약 |
| V9 | 표현력 공백의 명시적 계약화 - 지원 안 되는 조합(OR, 상수 피연산자 등)은 문서에 "안 된다"고 기재 | SV-C (조용한 오해 방지) |

## 4. 스코프 밖 메모

- 이 문서는 기준 도출까지만 담당한다. 전수조사 결과와 위반 목록은
  `../tasks/2026-07-10_chaining_api_audit.md`, 개선안·구현은 그 뒤 게이트에서 결정.
- 코드 생성(jOOQ류) 기반 타입 안전화는 VBA 도구 체인 부재로 1.0 스코프에서 제외.
