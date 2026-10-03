# UnivDB Library Reference (Consolidated)

> **For**: Claude Code agent context. Single-file reference replacing CheatSheet, Docs, Guide, Template.
> **Version**: v2.8.1

---

## 1. Architecture

- **Pattern**: Facade (`UnivDB` module) + Method Chaining
- **Core Objects**: `clsTable` (data container), `clsQuery` (CRUD), `clsReportBuilder` (ETL pipeline), `clsTableBuilder` (row-by-row assembly)
- **Data flow**: Sheet/Array → `RegisterTable` → in-memory `clsTable` → Query/Report/Join → output

---

## 2. UnivDB Facade (Entry Point)

```vb
UnivDB.InitDB                                          ' 엔진 초기화 (최초 1회)
UnivDB.RegisterTable "Name", Sheet1.Range("A1")        ' 시트→메모리 등록 (Smart Range)
UnivDB.GetTable("Name") -> clsTable                    ' 등록된 테이블 직접 참조
UnivDB.CreateTable() -> clsTable                       ' [Factory] 빈 clsTable 생성 (xlam 참조 시 New 불가 대응)
UnivDB.ClearDB                                         ' 메모리 정리

UnivDB.Query(Target) -> clsQuery                       ' Target: String(테이블명) | clsTable
UnivDB.Report(Target) -> clsReportBuilder              ' Target: String | clsTable
UnivDB.Join(Left, Right, Key, [JoinType], [Prefix], [RightKey]) -> clsTable
UnivDB.Upsert Target, Source, KeyCol                   ' 소스를 대상에 키 기준 병합 (clsTable.UpsertFrom 파사드)
UnivDB.Builder(KeyCol) -> clsTableBuilder              ' 빈 빌더 생성
UnivDB.EditTable(TableName, KeyCol) -> clsTableBuilder ' 기존 테이블 스키마 복제 빌더
```

---

## 3. clsQuery (Search & CRUD)

### Chainable Filters
```vb
.Where(Col, Val)              ' 인덱스 기반 동등검색 (최고속, 먼저 호출 권장)
.Filter(Col, Op, Val)         ' 조건 스캔 (>, <, >=, <=, =, <>, LIKE, NOT LIKE, IS, IS NOT)
.FilterAny(Array(Col, Op, Val), ...)  ' OR 그룹 (arm 중 하나라도 매칭 - 아래 참조)
```

> **FilterAny (교차 컬럼 OR)**: 한 호출이 하나의 OR 그룹이다 - 그룹 안 arm들은 OR,
> 그룹과 체인의 나머지는 AND (`(arm1 OR arm2) AND ...`). 각 arm은 `Array(컬럼, 연산자, 값)`
> 3요소이고 연산자는 `Filter`와 같은 화이트리스트를 쓴다. 동일 컬럼 arm을 나열하면
> SQL IN 의미론이 된다. arm 배열 하나(`FilterAny(vArmList)`)를 넘기는 동적 목록 형태도
> 허용. 전체 arm이 평가 전에 문장 수준 검증된다 (arm 하나가 틀리면 부분 축소 없이 9999).
>
> (v2.8.0 정정: 이전 문서의 `FilterIn`/`FilterNotIn`/`FilterLike`/`FilterNotLike`는
> 구현된 적 없는 유령 표기였다 - IN은 `FilterAny`의 동일 컬럼 arm으로 표현한다.)

### Sort / Limit / Project
```vb
.OrderBy(Col, [Desc])         ' Desc=True → 내림차순
.Top(Count)                   ' 상위 N행
.SelectColumns("Col1, Col2")  ' 컬럼 선택 (쉼표 구분 String)
```

### Join (물질화 장벽)
```vb
.Join(Right, LeftKey, [JoinType=jtInner], [RightPrefix], [RightKey]) -> clsQuery
```

> **물질화 장벽**: `Join`은 그 시점까지 스테이징된 조건/투영/정렬/Top을 **실행해 좌측
> 테이블을 확정**한 뒤, 조인 결과로 재소싱해 체인을 초기 상태에서 계속한다. Join 이전의
> `SelectColumns`/`OrderBy`/`Top`은 좌측에 소진되고 이후로 이월되지 않는다.
> `Right`는 테이블명(String) 또는 clsTable. 기본 JoinType은 **jtInner**(파사드 정합) —
> `clsReportBuilder.JoinTable`의 기본(jtLeft)과 다르니 주의. 키 컬럼 오타는 즉시 raise.

### Execute (Read)
```vb
.ToTable() -> clsTable        ' 결과를 clsTable로 물질화 (canonical 종결자)
.SelectTable() -> clsTable    ' ToTable과 동일 (호환 유지)
```

### Execute (Write) — Long 반환, 체이닝 마지막에 위치
```vb
.SetVal(Col, Val)             ' Insert/Update 값 스테이징 (canonical, 빌더 SetVal 정합)
.SetValue(Col, Val)           ' SetVal과 동일 (호환 유지)
.Insert() -> Long
.Update() -> Long             ' Where/Filter 조건 매칭 행 일괄 수정
.Delete() -> Long
```

### Example
```vb
' SELECT
UnivDB.Query("Emp") _
      .Where("Dept", "Sales") _
      .Filter("Salary", ">=", 5000) _
      .OrderBy("Salary", True) _
      .SelectColumns("Name, Role, Salary") _
      .SelectTable() _
      .DumpToSheet Sheet2.Range("A1"), True

' UPDATE
UnivDB.Query("Emp") _
      .Where("Name", "Bob") _
      .SetValue("Dept", "Finance") _
      .SetValue("Role", "Manager") _
      .Update
```

---

## 4. clsReportBuilder (ETL Pipeline)

> Copy-on-Write: 원본 데이터 보존. 메모리 상 가공 후 출력.

### Structure Modification
```vb
.AddColumns("Col1, Col2")                   ' 빈 컬럼 추가
.RemoveColumns("Col1, Col2")                ' 컬럼 삭제
.CreateKey(NewCol, SrcCol, Length)           ' Left() 기반 키 생성
```

### Join & Transform
```vb
.JoinTable(RefTable, Key, [JoinType], [Prefix], [RightKey])  ' 테이블 결합
.FillFromMap(TgtCol, SrcCol, Default)             ' VLOOKUP 효과
.ConvertDates(Col1, Col2, ...)                    ' 텍스트→날짜 변환
```

### Filter & Update
```vb
.Keep(Col, Op, Val)                                ' 조건 행 유지 (canonical, Exclude와 대칭)
.KeepAny(Array(Col, Op, Val), ...)                 ' OR 그룹 행 유지 (clsQuery.FilterAny 위임, 즉시 실행)
.Filter(Col, Op, Val)                              ' Keep과 동일 (호환 유지 - 주의: 즉시 실행)
.Exclude(Col, Op, Val)                             ' 조건 행 제거 (Remove)
.PatchIf(CondCol, Op, CondVal, TgtCol, NewVal)     ' 조건부 값 변경 (canonical, 빌더 PatchIf 인자 순서)
.UpdateValue(TgtCol, NewVal, CondCol, Op, CondVal) ' PatchIf와 동일 (호환 유지 - 인자 순서 주의)
```

> **실행 시점 주의**: `clsQuery.Filter`는 지연(후보 축소만, `ToTable()`에서 물질화)이지만
> ReportBuilder의 `Keep`/`KeepAny`/`Filter`/`Exclude`는 **즉시 실행**(작업 테이블을 그 자리에서 교체)이다.
> 신규 코드는 즉시 실행 어휘 `Keep`을 사용해 혼동을 피할 것.

### Output
```vb
.SelectColumns("Col1", "Col2", ...)                ' 최종 컬럼 순서 (canonical)
.SelectFinal("Col1", "Col2", ...)                  ' [Deprecated] SelectColumns 별칭 (호환 유지)
.DefineColumn(Col, [vWidth], [vFormat], [vAlign], [vHeaderAlias])
.BuildToSheet Sheet, "A1"                          ' 시트 렌더링 (서식 포함)
.ToTable() -> clsTable                             ' 작업 테이블 반환 (canonical, .Table과 동일)
```

### Example
```vb
UnivDB.Report("SalesData") _
      .Exclude("Status", "=", "Cancelled") _
      .Filter("Year", "=", 2024) _
      .UpdateValue("Grade", "VIP", "Amount", ">=", 10000) _
      .SelectFinal("Date", "Customer", "Amount", "Grade") _
      .DefineColumn("Date", vWidth:=12, vFormat:="yyyy-mm-dd", vAlign:="Center") _
      .DefineColumn("Amount", vWidth:=15, vFormat:="#,##0", vAlign:="Right") _
      .DefineColumn("Grade", vAlign:="Center", vHeaderAlias:="VIP Grade") _
      .BuildToSheet Sheet2, "A1"
```

---

## 5. clsTable (Core Container)

### I/O
```vb
.LoadFromSheet(Range)         ' 시트에서 로드
.LoadFromArray(Array)         ' 배열에서 로드
.DumpToSheet(Range, [bAutoFit], [eTextGuard], [bClearRegion])  ' 시트로 출력 (기본 옛 영역 선클리어)
.RawArray                     ' 내부 2D 배열 접근 (읽기 전용 권장)
```

> **DumpToSheet 클리어 계약 (기본 클리어)**: `DumpToSheet`는 기본으로(`bClearRegion:=True`)
> 쓰기 전 시작셀의 옛 `CurrentRegion`을 선클리어한 뒤 새 데이터를 덮어쓴다. 축소 재덤프에서
> 옛 초과행이 잔존해 하위 `CurrentRegion` 읽기에 유령행이 재유입되던 silent corruption을
> 기본에서 차단한다(빈 결과 덤프도 영역을 비운다). 인접 보조열·수식을 보존해야 하는 외과적
> 쓰기만 `bClearRegion:=False`로 옵트아웃한다. `clsReportBuilder.BuildToSheet`는 자체적으로
> 영역 위생을 관리하므로 내부에서 `bClearRegion:=False`로 옵트아웃한다.
> 상세: `src/docs/tasks/2026-07-14_dumptosheet_no_clear_footgun.md`

### Data Operations

변형(mutation) 메서드는 전부 자기 자신(`clsTable`)을 반환해 체이닝이 가능하다
(in-place 변형이며 세대(Generation)는 규칙대로 증가한다):

```vb
.UpsertFrom(SourceTable, KeyCol)    ' 키 기준 Insert or Update (Merge)
.GroupBy(GroupCols, AggList)         ' 집계: "Col:SUM, Col:AVG, Col:COUNT, Col:MAX, Col:MIN"
                                     '   함수 생략 시 COUNT 기본 ("Col" == "Col:COUNT")
.Compute(TgtCol, Op1, Op, Op2)      ' 컬럼/상수 연산: +, -, *, /, & - 피연산자가 String이면
                                     '   컬럼명(부재 시 9999), 그 외 타입이면 숫자 상수
                                     '   (예: .Compute("Tax", "Amt", "*", 0.1))
.OrderBy(Col, [bDescending])        ' 정렬 (in-place, Me 반환)
.AddColumn(Col, [Init])             ' 컬럼 추가 - 멱등 ensure (이미 있으면 무시)
.AddColumns("B, C") / ("B", "C")    ' 복수 추가 - 개별 인자/csv/Array 허용 (멱등 ensure)
.RemoveColumn(Col)                  ' 컬럼 삭제 - 멱등 drop (없으면 무시)
.RemoveColumns("B, C") / ("B", "C") ' 복수 삭제 - 개별 인자/csv/Array 허용 (멱등 drop)
.CopyColumn(Src, Tgt)               ' 고속 컬럼 복사 (Src 오타는 raise)
.UpdateBatch(SetCol, SetVal, WhereCol, WhereVal)  ' 동등 조건 일괄 변경
.DeleteBatch(WhereCol, WhereVal)    ' 동등 조건 일괄 삭제 (0건 매칭은 무동작)

' 체이닝 예:
Set tTop = T.GroupBy("Dept", "Amt:SUM").OrderBy("Amt_SUM", True)
T.AddColumn("Bonus", 0).Compute("Total", "Amt", "+", "Bonus").OrderBy "Total", True
```

### Properties
```vb
.Headers             ' Dictionary 사본 — 컬럼명→인덱스 매핑, 수정해도 원본에 영향 없음
.HasColumn("Name")   ' 존재 여부, 대소문자/앞뒤 공백 무시
.ColumnIndex("Name") ' 열 번호, 미존재 시 9999
.data       ' 2D Array (내부 데이터)
```

---

## 6. Value & Key Semantics (확정 계약)

> "조용히 틀린 결과"를 막기 위한 값/키 의미론 계약 (비결정 D, 2026-07-10 확정).
> 인덱스/필터/정렬/집계 전 표면에서 동일하게 적용된다. 회귀 핀: `TestNullTypeSemantics.bas`

### 6.1 인덱스 키 정규화 (`modIndexKey.NormalizeIndexKey`)

Where/Join/GroupBy 키는 전부 단일 정규화 계약을 탄다:

| 값 | 키 |
|---|---|
| Empty (빈 셀) | `""` (빈 문자열 키) |
| Date | ISO canonical: `yyyy-mm-dd` (시간 성분 0) / `yyyy-mm-dd hh:nn:ss` — locale 무관 |
| Error (#N/A 등) | 키 없음 — 해당 행이 인덱스/그룹핑에서 skip |
| 그 외 | `CStr` 문자열화 — 숫자 `42`와 문자 `"42"`는 **같은 키로 병합** |

- 키 비교는 대소문자 무시 (`vbTextCompare`)
- **Unicode 한계 (2026-10-03 실측)**: 기본 Dictionary 및 복합 GroupBy의 Dictionary는
  `StrComp(..., vbTextCompare)`와 모든 Unicode 입력에서 동치가 아니다.
  이 Windows/Excel 환경에서 U+00E9와 U+0065 U+0301, `"ab"`와 `"a" & ChrW(0) & "b"`는
  StrComp에서는 같지만 Dictionary 키에서는 분리됐다. Core의 텍스트 해시는 Windows sort key를
  사용해 자기 비교기와의 불일치를 수정했으며, 기본 백엔드/복합 키의 Unicode 통일은 미해결이다.
  백엔드 간 동치 보장은 이 입력들에 적용되지 않는다. 자동 백엔드 전환은 없다.
  Core는 [LCMapStringW sort key](https://learn.microsoft.com/en-us/windows/win32/api/winnls/nf-winnls-lcmapstringw)의
  바이트열을 해싱한다. 기존 2만 행 벤치의 3회 중앙값은 수정 전 23/25ms에서 수정 후 102/103ms로
  증가했다(저중복/고중복, 동일 세션 실측). 기본 Dict는 49/32ms에서 50/32ms였다.
- 복합 키(GroupBy 다중 컬럼)는 길이 접두 인코딩 `len:value|len:value|...` — 길이가 파싱 권위라 값에 구분자(`|`, `:`)가 있어도 충돌 없음
- GroupBy 결과의 그룹 컬럼 값은 정규화 키 문자열 그대로 렌더된다 (Date면 ISO 표기)

**시트 로드의 날짜 경계**: ISO 규칙은 Variant의 실제 타입이 `vbDate`일 때만 적용된다.
`LoadFromSheet`는 [Range.Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2)를
사용하므로 날짜/통화 서식의 셀도 Double로 읽는다. 날짜 서식이나 컬럼명으로 타입을 추론하지 않는다.
이 경로의 날짜 인덱스 키는 숫자 일련값의 `CStr`이고, Date/ISO 키와 자동 병합되지 않는다.

Excel의 서식 기반 Date 타입이 필요하면 호출자가 명시적으로
`T.LoadFromArray ws.Range("A1:D10").Value`를 사용할 수 있다.
`Range.Value`는 Currency 타입도 사용하므로 숫자 정밀도 정책까지 함께 선택하는 입력 경계다.
서식에 의존하지 않으려면 업무 스키마에 따라 Date Variant 배열을 구성해 `LoadFromArray`에 전달한다.

### 6.2 NULL/타입 계약 (Empty / Error)

- **Empty = NULL**: 숫자 비교(`Filter` `>`, `<`, `=` 등) 불참, 모든 집계 불참. 문자 경로(IS/LIKE)와 인덱스 키에서만 `""`로 참여.
- **Error = skip**: 비교/정렬/집계/필터 어디서도 값으로 참여하지 않는다.
  - Filter(숫자/문자): Error 행 제외 — `IS NOT`/`NOT LIKE` 등 부정 연산 결과에도 안 들어온다
  - OrderBy: 크래시 없이 결정적으로 맨 뒤 배치 (행은 보존)
  - GroupBy: 그룹 키면 행 skip, 집계 컬럼이면 해당 셀 불참

### 6.3 집계 규칙 (GroupBy)

| 값 | COUNT | SUM | AVG 분모 | MIN/MAX |
|---|---|---|---|---|
| Empty | 불참 | 불참 | 불참 | 불참 |
| Error | 불참 | 불참 | 불참 | 불참 |
| 숫자 (문자 `"42"` 포함, IsNumeric) | 포함 | 가산 | 포함 | 참여 |
| 비숫자 실값 (Date, 일반 문자) | 포함 | 불참 | 불참 | 참여 (Variant 비교) |

`AVG = SUM / (숫자로 가산된 셀 수)` — 비숫자 실값이 섞여도 평균이 희석되지 않는다.

### 6.4 OrderBy 비교 규칙 (`clsTable.CompareValues`)

비교 상대와 무관한 타입 순서를 사용한다. 오름차순은 다음 순서다.

1. **숫자와 Empty** → CDbl 수치 비교 (Empty는 0)
2. **Date** → CDbl 직렬값 비교
3. **Boolean** → False/True의 CDbl 값 비교
4. **String** → `vbTextCompare` 텍스트 비교 (숫자 문자열도 String 그룹)
5. **Null**
6. **Error**

내림차순은 일반 값의 순서를 뒤집으며 **Error는 양방향 모두 마지막**이다.
기존의 "한쪽이 String이면 문자 비교" 정책은 비교 순환을 만들므로 폐기했다.
따라서 숫자 `2`, 숫자 `10`, 문자열 `"15"`의 오름차순은 언제나 `2, 10, "15"`다.
동등 키 사이의 안정 정렬은 보장하지 않는다.

쿼리는 조건 후보의 원본 rowID를 정렬한 뒤 Top과 열 선택을 적용한다.
정렬 키를 출력 열에서 제외할 수 있고, 원본 데이터/Generation은 바뀌지 않는다.
0행 결과 정렬은 헤더만 유지한다. 존재하지 않는 정렬 열은 9999로 거부한다.

### 6.5 연산자 화이트리스트와 fail-fast (`modOperators`)

연산자/집계함수 문자열을 받는 모든 표면은 `modOperators`의 단일 정규화 지점을 경유하며,
**미지 문자열은 즉시 `Err.Raise 9999`** 한다 (조용한 0행/무동작/빈 컬럼 없음 — 컬럼명
오타의 9999와 오류 처리 대칭):

| 표면 | 허용 목록 |
|---|---|
| `Filter` / `FilterAny` / `Keep` / `KeepAny` / `Exclude` / `UpdateValue` 연산자 | `=` `<>` `>` `<` `>=` `<=` `IS` `IS NOT` `LIKE` `NOT LIKE` (대소문자/공백 무관) |
| `GroupBy` 집계함수 | `SUM` `AVG` `COUNT` `MIN` `MAX` |
| `Compute` 연산자 | `+` `-` `*` `/` `&` |
| `PatchIf` 연산자 | `=` `<>` `>` `>=` `<` `<=` |

**문장 수준 vs 데이터 수준 오류의 구분**: 개발자가 쓴 문장의 오류(미지 연산자, 컬럼명
오타, 수치 연산자에 비숫자 비교값)는 raise. 데이터 셀 수준의 문제(Empty/Error/비숫자
셀)는 6.2~6.3의 NULL/skip 계약대로 해당 셀만 불참 — 크래시도, 조용한 오염도 없다.

v2.8.0에서 추가된 문장 수준 raise: `FilterAny`의 0개 arm/3요소 아닌 arm/비스칼라 arm 요소
(전체 arm이 평가 앞에서 검증 - 부분 축소 없음), `Compute`의 비String 비숫자 피연산자
(Empty/Null/Date/배열/객체)와 상수 0 나눗셈 (컬럼 셀의 0 나눗셈은 기존 데이터 수준 관용 유지).

`Where`/`Filter`의 컬럼과 `Filter`의 비교값 검증은 후보 0행에서도 생략하지 않는다.
`Update`는 모든 SET 컬럼을 첫 쓰기 전에 검증하며, 컬럼 오타로 일부 값만 변경되지 않는다.
이는 실행 중 발생하는 모든 오류를 롤백하는 트랜잭션 보장은 아니다.

값만 바꾸면 rowID Generation은 유지되지만, 해당 열의 캐시에서 폐기된 인덱스 객체는
`EnsureCurrent` 및 조회에서 9998로 실패한다. 인덱스는 테이블에서 다시 취득해야 한다.
`DbColumnReader` 역시 값 사본이지만 `EnsureCurrent(T)`는 구조 세대와 해당 열의 값 버전을
함께 검사한다. `Value`/`UpdateBatch`/`CopyColumn`/`Compute`가 그 열을 변경하면 9998로 실패한다.
다른 열만 바뀌면 통과한다. `Values`/`ValueAt`은 자동 갱신되지 않는 사본 읽기이므로,
현재 값이 필요한 호출자는 `EnsureCurrent`로 확인하거나 reader를 다시 취득한다.

이미 얻은 RowList/RowSet/커서는 기존 선택의 스냅샷이다. 세대 검사는 행 좌표의 유효성만
검증하며 변경된 값으로 조건을 다시 평가하지 않는다.

### 6.6 표현력 공백 (명시적 비지원)

- **조건 결합: 체인은 AND, OR은 `FilterAny` 그룹 단위로만** (v2.8.0에서 교차 컬럼 OR
  지원). 표현 가능한 형태는 정확히 `(arm1 OR arm2 OR ...) AND (다음 그룹/조건) AND ...` -
  즉 AND of OR-groups. **OR 그룹 안에 AND를 중첩하는 형태(`(A AND B) OR C`)는 여전히 표현
  불가**가 계약이다 (분배법칙으로 AND-of-ORs로 변환하거나 쿼리를 분리할 것). 동일 컬럼
  값 목록 매칭(IN)은 `FilterAny`의 동일 컬럼 arm으로 표현한다.
- **`Compute` 상수는 숫자만**: 피연산자 구분 규칙이 "String=컬럼명, 그 외=숫자 상수"라
  **문자열 리터럴 상수는 표현 불가** (`Compute("T", "A", "&", "-")`는 컬럼 `-`를 찾다
  raise). 문자열 결합에 리터럴이 필요하면 `AddColumn`으로 상수 컬럼을 만들어 쓸 것.
- **`AddColumn`은 멱등 ensure / `RemoveColumn`은 멱등 drop**: 존재/부재 시 조용히
  무시가 계약이다 (컬럼명 오타 검출이 필요하면 `Headers.exists`로 확인).

### 6.7 어휘 지도 (Vocabulary Map)

같은 개념의 canonical 이름과 표면별 대응. 레거시 이름은 전부 영구 유지된다.
빈 칸은 해당 표면에 그 개념이 없음.

| 개념 | canonical | clsQuery | clsReportBuilder | clsTableBuilder | clsTable |
|---|---|---|---|---|---|
| 종결 (테이블 물질화) | `ToTable` | ToTable = SelectTable | ToTable = Table(Get) | ToTable | (자신이 테이블) |
| 지연 필터 (후보 축소) | `Filter`/`Where` | Filter, Where | | | |
| OR 그룹 (교차 컬럼/IN) | `FilterAny` | FilterAny (지연) | KeepAny (즉시) | | |
| 즉시 필터 (행 유지/제거) | `Keep`/`Exclude` | | Keep = Filter / Exclude | | |
| 값 스테이징 | `SetVal` | SetVal = SetValue | | SetVal (Val은 레거시 혼성) | |
| 조건부 패치 | `PatchIf` (cond,op,val,tgt,new) | | PatchIf = UpdateValue(인자 순서 상이) | PatchIf | |
| 직접 셀 쓰기 | `Value` (Let) | | | | Value(row, col) = v |
| 행 제거 | 개념별 상이 | Delete() -> Long (스테이지드 CRUD) | Exclude (술어 즉시) | | DeleteBatch (동등 일괄) |
| 컬럼 추가/삭제 | `AddColumns`/`RemoveColumns` | | AddColumns/RemoveColumns (csv 전용 - 레거시 규약) | (자동 확장) | AddColumns/RemoveColumns (개별/csv/Array) + 단수형 |
| 조인 | `Join` | Join (기본 jtInner, 물질화 장벽) | JoinTable (기본 jtLeft) | | (엔진: modEngine) |
| 진입점 | 명사형 | UnivDB.Query | UnivDB.Report | UnivDB.Builder / EditTable(동사구 - 별칭 없이 유지) | UnivDB.GetTable/CreateTable |

- 행 제거 3종은 **서로 다른 연산**이다 (스테이지드 CRUD / 술어 즉시 / 동등 일괄) -
  이름을 통일하지 않는 것이 의도된 결정 (같은 이름 = 같은 의미론 원칙).
- 목록 인자 규약: canonical은 "ParamArray + 개별 인자/csv 한 줄/Array 모두 허용"
  (clsTable.SelectColumns/AddColumns/RemoveColumns). RB의 csv 전용은 레거시.

---

## 7. clsTableBuilder (Dynamic Row Assembly)

```vb
Dim B As clsTableBuilder
Set B = UnivDB.Builder("PROD_ID")

' canonical: SetVal(순수 setter, 빌더 반환) / GetVal(getter, 미지 컬럼 raise)
' 체인 마지막 호출을 괄호 없이 쓰면 Call 불필요:
B.Row("P001").SetVal("Name", "Mouse").SetVal("Qty", 10).SetVal "Price", 15000
' 또는 전체를 Call로 감싼다:
Call B.Row("P002").SetVal("Name", "Keyboard").SetVal("Qty", 0).PatchIf("Qty", "=", 0, "Status", "SoldOut")

Debug.Print B.GetVal("Qty")   ' 커서 행 값 읽기 - 미지 컬럼이면 Err.Raise 9999

Dim tNew As clsTable
Set tNew = B.ToTable()

' [Legacy] Val(col[, value])은 getter/setter 혼성 - 호환 유지, 신규 코드는 SetVal/GetVal
```

> **VBA 체인 구문 계약**: 문(statement) 위치에서 괄호 붙은 다인자 호출로 체인을 끝내면
> VBA 구문 오류다. (1) 마지막 호출만 괄호 없이 쓰거나 (2) 전체를 `Call`로 감싼다.
> 이는 모든 빌더 체인에 공통인 VBA 제약이다.

---

## 8. JoinType Enum

| Enum | Value | Description |
|------|-------|-------------|
| `jtInner` | 1 | 교집합 |
| `jtLeft` | 2 | Left Outer Join |
| `jtRight` | 3 | Right Outer Join |
| `jtFull` | 4 | Full Outer Join |
| `jtLeftAnti` | 5 | 왼쪽 차집합 (B에 없는 A) |
| `jtRightAnti` | 6 | 오른쪽 차집합 (A에 없는 B) |
| `jtLeftSemi` | 7 | 왼쪽 존재확인 (B에 있는 A, 오른쪽 속성 제외) |
| `jtRightSemi` | 8 | 오른쪽 존재확인 |

---

## 9. Code Templates (Standard Patterns)

> **에이전트는 새 모듈 작성 시 아래 패턴을 기준으로 삼아야 합니다.**

### 9.1 clsTable 생성 — 3가지 경로

```vb
' ── (A) 시트에서 직접 생성 (RegisterTable 불필요, 일회성 사용) ──
Dim tLocal As New clsTable
tLocal.LoadFromSheet Sheet1.Range("A1").CurrentRegion

' ── (B) 전역 저장소에서 꺼내기 (InitDB + RegisterTable 선행 필수) ──
Dim tEmp As clsTable
Set tEmp = UnivDB.GetTable("Employees")

' ── (C) Query 결과로 생성 (필터링된 부분집합) ──
Dim tFiltered As clsTable
Set tFiltered = UnivDB.Query("Employees") _
                      .Where("Dept", "Sales") _
                      .SelectTable()
```

### 9.2 Cascade Join (★ 핵심 패턴)

```vb
' tMain 객체를 덮어쓰며 연쇄적으로 참조 테이블을 확장합니다.
' UnivDB.Join 시그니처:
'   UnivDB.Join(vLeft, vRight, sLeftKey, [JoinType], [RightPrefix], [RightKey]) -> clsTable
'
'   vLeft      : clsTable 객체 또는 테이블명(String)
'   vRight     : 테이블명(String) — RegisterTable로 등록된 이름
'   sLeftKey   : Left 테이블의 조인 키 컬럼명
'   JoinType   : jtLeft(2) 권장 (미매칭 시 빈칸 유지)
'   RightPrefix: 컬럼명 충돌 방지 접두사 (예: "Proj_", "Client_")
'   RightKey   : 오른쪽 테이블의 조인 키 (생략 시 sLeftKey와 동일)

Set tMain = UnivDB.Join(tMain, "tbl_Project", "Ref_ProjectID", jtLeft, "Proj_")
Set tMain = UnivDB.Join(tMain, "tbl_Client",  "Ref_ClientID",  jtLeft, "Client_")
' → tMain.headers에 "Proj_Name", "Client_Name" 등 접두사 컬럼이 추가됨
```

### 9.3 Factory Function 패턴 (modPresetFactory 기준)

```vb
' ── 팩토리 함수: clsTable을 받아 프리셋별 조인을 적용하고 반환 ──
Public Function Build_Master_Table(ByVal presetName As String, _
                                   ByVal tMain As clsTable) As clsTable
    Select Case presetName
        Case "통합계약대장"
            Set tMain = UnivDB.Join(tMain, "tbl_Project", "Ref_ProjectID", jtLeft, "Proj_")
            Set tMain = UnivDB.Join(tMain, "tbl_Client",  "Ref_ClientID",  jtLeft, "Client_")
        Case "검수현황"
            Set tMain = UnivDB.Join(tMain, "tbl_Inspect", "Ref_InspectID", jtLeft, "Insp_")
        Case Else
            ' 조인 불필요 — tMain을 그대로 통과
    End Select
    
    Set Build_Master_Table = tMain
End Function
```

### 9.4 End-to-End Pipeline (Broker → Factory → Report)

```vb
' ── modDashboardBroker.Action_Refresh_Dashboard 내부 흐름 ──
Sub Action_Refresh_Dashboard(presetName As String, tblSource As ListObject)
    On Error GoTo ErrHandler
    
    UnivDB.InitDB                                 ' (0) DB 엔진 초기화
    ' ... RegisterTable 등 참조 테이블 등록 ...
    
    Dim oTable As New clsTable
    oTable.LoadFromSheet tblSource.Range          ' (1) 시트 → 메모리

    ' --- Master DB Join Pipeline ---
    Set oTable = modPresetFactory.Build_Master_Table(presetName, oTable)  ' (2) 팩토리 조인
    ' ---

    Set oBuilder = UnivDB.Report(oTable)          ' (3) 리포트 빌더 시작
    ' ... .Exclude / .Filter / .SelectFinal / .BuildToSheet ...

Finally:
    UnivDB.ClearDB                                ' ★ 반드시 정리 — clsTable 메모리 해제
    Exit Sub
ErrHandler:
    modLogger.LogError Err, "Action_Refresh_Dashboard"
    Resume Finally
End Sub
```

> **Safety Note**: 조인으로 추가된 Prefix 컬럼들은 `tSource.headers.exists` 체크에 의해
> Upsert 시 자동으로 무시됩니다 (Read-Only 효과). Upsert 관련 코드 수정 금지.

---

## 10. Coding Conventions

- **DB Lifecycle**: `InitDB`를 호출하는 서브루틴은 반드시 `Finally` 블록에서 `ClearDB`를 호출해야 합니다. 서브루틴 밖에서 clsTable이 살아있으면 안 됩니다. (9.4 패턴 참조)
- **변수 접두사**: `t` = clsTable, `o` = Object, `s` = String, `b` = Boolean, `v` = Variant
- **함수명**: PascalCase (`Build_Master_Table`)
- **에러 핸들링**: `On Error GoTo ErrHandler` + `modLogger.LogError` + `Resume Finally` 패턴 사용
- **CRUD 반환**: `Long` (처리 행 수) → `Set` 사용 안함, 체이닝 끝에 위치
- **Direct Object**: `RegisterTable` 없이 `clsTable` 객체를 `Query()`/`Report()`에 직접 전달 가능
