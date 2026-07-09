# UnivDB Library Reference (Consolidated)

> **For**: Claude Code agent context. Single-file reference replacing CheatSheet, Docs, Guide, Template.
> **Version**: v2.6.0

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

### Execute (Read)
```vb
.SelectTable() -> clsTable    ' 결과를 clsTable로 반환
```

### Execute (Write) — Long 반환, 체이닝 마지막에 위치
```vb
.SetValue(Col, Val)           ' Insert/Update 값 설정 (체이닝 가능)
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
.Filter(Col, Op, Val)                              ' 조건 행 유지 (Keep)
.Exclude(Col, Op, Val)                             ' 조건 행 제거 (Remove)
.UpdateValue(TgtCol, NewVal, CondCol, Op, CondVal) ' 조건부 값 변경 (IF 효과)
```

### Output
```vb
.SelectFinal("Col1", "Col2", ...)                  ' 최종 컬럼 순서
.DefineColumn(Col, [vWidth], [vFormat], [vAlign], [vHeaderAlias])
.BuildToSheet Sheet, "A1"                          ' 시트 렌더링 (서식 포함)
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
```vb
.UpsertFrom(SourceTable, KeyCol)    ' 키 기준 Insert or Update (Merge)
.GroupBy(GroupCols, AggList)         ' 집계: "Col:SUM, Col:AVG, Col:COUNT, Col:MAX, Col:MIN"
.Compute(TgtCol, Col1, Op, Col2)    ' 컬럼 간 연산: +, -, *, /, &
.OrderBy(Col, [bDescending])        ' 정렬
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

---

## 7. clsTableBuilder (Dynamic Row Assembly)

```vb
Dim B As clsTableBuilder
Set B = UnivDB.Builder("PROD_ID")

' Call 필수 (체이닝 시)
Call B.Row("P001").Val("Name", "Mouse").Val("Qty", 10).Val("Price", 15000)
Call B.Row("P002").Val("Name", "Keyboard").Val("Qty", 0).PatchIf("Qty", "=", 0, "Status", "SoldOut")

Dim tNew As clsTable
Set tNew = B.ToTable()
```

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
