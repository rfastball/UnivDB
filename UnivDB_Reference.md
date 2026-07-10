# UnivDB Library Reference (Consolidated)

> **For**: Claude Code agent context. Single-file reference replacing CheatSheet, Docs, Guide, Template.
> **Version**: v2.7.2

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
UnivDB.ClearDB                                         ' 메모리 정리

UnivDB.Query(Target) -> clsQuery                       ' Target: String(테이블명) | clsTable
UnivDB.Report(Target) -> clsReportBuilder              ' Target: String | clsTable
UnivDB.Join(Left, Right, Key, [JoinType], [Prefix], [RightKey]) -> clsTable
UnivDB.Builder(KeyCol) -> clsTableBuilder              ' 빈 빌더 생성
UnivDB.EditTable(TableName, KeyCol) -> clsTableBuilder ' 기존 테이블 스키마 복제 빌더
```

---

## 3. clsQuery (Search & CRUD)

### Chainable Filters
```vb
.Where(Col, Val)              ' 인덱스 기반 동등검색 (최고속, 먼저 호출 권장)
.Filter(Col, Op, Val)         ' 조건 스캔 (>, <, >=, <=, =, <>, LIKE, NOT LIKE, IS, IS NOT)
.FilterIn(Col, Val1, ...)     ' SQL IN
.FilterNotIn(Col, Val1, ...)  ' SQL NOT IN
.FilterLike(Col, Pat1, ...)   ' 패턴 OR 매칭
.FilterNotLike(Col, Pat...)   ' 패턴 OR 제외
```

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
.Filter(Col, Op, Val)                              ' Keep과 동일 (호환 유지 - 주의: 즉시 실행)
.Exclude(Col, Op, Val)                             ' 조건 행 제거 (Remove)
.PatchIf(CondCol, Op, CondVal, TgtCol, NewVal)     ' 조건부 값 변경 (canonical, 빌더 PatchIf 인자 순서)
.UpdateValue(TgtCol, NewVal, CondCol, Op, CondVal) ' PatchIf와 동일 (호환 유지 - 인자 순서 주의)
```

> **실행 시점 주의**: `clsQuery.Filter`는 지연(후보 축소만, `ToTable()`에서 물질화)이지만
> ReportBuilder의 `Keep`/`Filter`/`Exclude`는 **즉시 실행**(작업 테이블을 그 자리에서 교체)이다.
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
.DumpToSheet(Range, [bAutoFit])  ' 시트로 출력
.RawArray                     ' 내부 2D 배열 접근 (읽기 전용 권장)
```

### Data Operations

변형(mutation) 메서드는 전부 자기 자신(`clsTable`)을 반환해 체이닝이 가능하다
(in-place 변형이며 세대(Generation)는 규칙대로 증가한다):

```vb
.UpsertFrom(SourceTable, KeyCol)    ' 키 기준 Insert or Update (Merge)
.GroupBy(GroupCols, AggList)         ' 집계: "Col:SUM, Col:AVG, Col:COUNT, Col:MAX, Col:MIN"
                                     '   함수 생략 시 COUNT 기본 ("Col" == "Col:COUNT")
.Compute(TgtCol, Col1, Op, Col2)    ' 컬럼 간 연산: +, -, *, /, & (피연산자는 컬럼명만,
                                     '   상수 리터럴 미지원)
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
.headers    ' Dictionary — 컬럼명→인덱스 매핑
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
- 복합 키(GroupBy 다중 컬럼)는 길이 접두 인코딩 `len:value|len:value|...` — 길이가 파싱 권위라 값에 구분자(`|`, `:`)가 있어도 충돌 없음
- GroupBy 결과의 그룹 컬럼 값은 정규화 키 문자열 그대로 렌더된다 (Date면 ISO 표기)

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

우선순위 순서대로:

1. **Error** → 맨 뒤 (결정적, locale 무관)
2. **Date + Date** → 직렬값(CDbl) 수치 비교 — 시간순 보장, locale 무관
3. **한쪽이 String** → 텍스트 비교(`vbTextCompare`) — 숫자 `9` vs 문자 `"10"`이면 `"10"`이 앞선다
4. **둘 다 IsNumeric** → CDbl 수치 비교 — Empty는 0 취급
5. **그 외** → 텍스트 fallback

### 6.5 연산자 화이트리스트와 fail-fast (`modOperators`)

연산자/집계함수 문자열을 받는 모든 표면은 `modOperators`의 단일 정규화 지점을 경유하며,
**미지 문자열은 즉시 `Err.Raise 9999`** 한다 (조용한 0행/무동작/빈 컬럼 없음 — 컬럼명
오타의 9999와 오류 처리 대칭):

| 표면 | 허용 목록 |
|---|---|
| `Filter` / `Keep` / `Exclude` / `UpdateValue` 연산자 | `=` `<>` `>` `<` `>=` `<=` `IS` `IS NOT` `LIKE` `NOT LIKE` (대소문자/공백 무관) |
| `GroupBy` 집계함수 | `SUM` `AVG` `COUNT` `MIN` `MAX` |
| `Compute` 연산자 | `+` `-` `*` `/` `&` |
| `PatchIf` 연산자 | `=` `<>` `>` `>=` `<` `<=` |

**문장 수준 vs 데이터 수준 오류의 구분**: 개발자가 쓴 문장의 오류(미지 연산자, 컬럼명
오타, 수치 연산자에 비숫자 비교값)는 raise. 데이터 셀 수준의 문제(Empty/Error/비숫자
셀)는 6.2~6.3의 NULL/skip 계약대로 해당 셀만 불참 — 크래시도, 조용한 오염도 없다.

### 6.6 표현력 공백 (명시적 비지원)

- **조건 결합은 AND 전용**: `Where`/`Filter` 체인은 항상 논리곱. 교차 컬럼 OR은 표현
  불가. 동일 컬럼의 OR(값 목록 매칭)은 `FilterIn(Col, v1, v2, ...)`이 커버한다.
- **`Compute`의 피연산자는 컬럼명만**: 상수 리터럴 미지원 (`Compute("T", "A", "+", "3")`은
  컬럼 `3`을 찾다 raise).
- **`AddColumn`은 멱등 ensure / `RemoveColumn`은 멱등 drop**: 존재/부재 시 조용히
  무시가 계약이다 (컬럼명 오타 검출이 필요하면 `Headers.exists`로 확인).

### 6.7 어휘 지도 (Vocabulary Map)

같은 개념의 canonical 이름과 표면별 대응. 레거시 이름은 전부 영구 유지된다.
빈 칸은 해당 표면에 그 개념이 없음.

| 개념 | canonical | clsQuery | clsReportBuilder | clsTableBuilder | clsTable |
|---|---|---|---|---|---|
| 종결 (테이블 물질화) | `ToTable` | ToTable = SelectTable | ToTable = Table(Get) | ToTable | (자신이 테이블) |
| 지연 필터 (후보 축소) | `Filter`/`Where` | Filter, Where, FilterIn 등 | | | |
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
