# UnivDB 101 - 처음 시작하는 사용자 가이드

> **대상**: UnivDB를 참조 라이브러리로 사용해 업무 매크로를 작성하는 VBA 사용자
> **기준 버전**: v2.8.0 / 전체 계약 명세는 `UnivDB_Reference.md` 참조

---

## 1. UnivDB란

UnivDB는 Excel 배치 작업(1만~10만 행 규모의 적재 → 변환 → 보고)을 위한 **인메모리 DB
라이브러리**입니다. 시트 데이터를 메모리로 올려 SQL처럼 조회·조인·집계한 뒤 시트로
출력합니다. 셀 단위 루프 대신 배열 기반으로 동작해 수만 행에서도 빠릅니다.

설계 철학은 **fail-fast**: 컬럼명 오타, 잘못된 연산자 등 코드의 실수는 조용히 0행을
돌려주지 않고 즉시 오류(`Err 9999`)를 냅니다. "조용히 틀린 결과"가 없도록 만든
의도된 동작입니다.

---

## 2. 설치 (xlam 참조 추가)

1. 배포본 `UnivDB.xlam` 파일을 받아 적당한 고정 경로에 둡니다.
2. 내 매크로 워크북을 열고 VBE(Alt+F11) → **도구 > 참조** → **찾아보기**로
   `UnivDB.xlam`을 선택해 체크합니다.
3. 이제 내 모듈에서 `UnivDB.InitDB`, `Dim t As clsTable` 등을 바로 쓸 수 있습니다.

> **중요**: 참조 방식에서는 `New clsTable`을 직접 쓸 수 없습니다.
> 빈 테이블이 필요하면 반드시 팩토리를 사용하세요:
>
> ```vb
> Dim t As clsTable
> Set t = UnivDB.CreateTable()   ' O
> ' Dim t As New clsTable        ' X - 참조 프로젝트 밖에서는 컴파일 오류
> ```

---

## 3. 5분 시작

시트의 표를 등록하고, 조건 조회 결과를 다른 시트에 출력하는 완결 매크로입니다.

```vb
Sub QuickStart()
    On Error GoTo ErrHandler

    ' (1) 엔진 초기화 + 시트 데이터 등록 (헤더 포함 표, 시작 셀만 주면 범위 자동 인식)
    UnivDB.InitDB
    UnivDB.RegisterTable "Emp", Sheet1.Range("A1")

    ' (2) 조회: 인덱스 동등검색(Where) -> 조건 스캔(Filter) -> 정렬 -> 컬럼 선택 -> 출력
    UnivDB.Query("Emp") _
          .Where("Dept", "Sales") _
          .Filter("Salary", ">=", 5000) _
          .OrderBy("Salary", True) _
          .SelectColumns("Name, Role, Salary") _
          .ToTable() _
          .DumpToSheet Sheet2.Range("A1"), True

Finally:
    ' (3) 반드시 정리 - InitDB를 호출한 프로시저는 ClearDB로 끝난다
    UnivDB.ClearDB
    Exit Sub
ErrHandler:
    MsgBox "오류 " & Err.Number & ": " & Err.Description
    Resume Finally
End Sub
```

이 한 개의 매크로에 UnivDB 사용의 뼈대가 전부 들어 있습니다:
**InitDB → RegisterTable → 작업 → 종결자 → ClearDB**.

---

## 4. 핵심 개념: 테이블 하나, 작업 표면 넷

모든 데이터는 `clsTable`(테이블 객체)에 담깁니다. 이 테이블을 다루는 진입점이 4개
있고, 각각 실행 방식이 다릅니다:

| 진입점 | 반환 | 성격 | 언제 쓰나 |
|---|---|---|---|
| `UnivDB.Query(대상)` | clsQuery | **지연 실행** - 조건을 쌓다가 종결자(`ToTable`/`Update` 등)에서 한 번에 실행 | 조회, 행 단위 CRUD |
| `UnivDB.Report(대상)` | clsReportBuilder | **즉시 실행** ETL - 각 단계가 그 자리에서 실행, 원본은 Copy-on-Write로 보호 | 보고서 가공 후 시트 출력 |
| `UnivDB.Builder(키)` / `EditTable(이름, 키)` | clsTableBuilder | 커서 기반 **행 조립** | 행을 하나씩 만들어 병합 |
| `clsTable` 자체 메서드 | clsTable (자기 자신) | **in-place 변형** - 원본을 직접 바꿈 | 집계, 컬럼 연산, 정렬 |

`대상` 자리에는 등록된 테이블명(String)과 clsTable 객체를 혼용할 수 있습니다.

### clsTable 직접 가공 맛보기 (집계·컬럼 연산)

```vb
Dim tSales As clsTable, tSum As clsTable
Set tSales = UnivDB.GetTable("Sales")

' 파생 컬럼: Tax = Amt * 0.1  (피연산자가 String이면 컬럼명, 숫자면 상수)
tSales.Compute "Tax", "Amt", "*", 0.1

' 부서별 집계 후 내림차순 정렬 (결과 컬럼명은 "Amt_SUM"처럼 자동 명명)
Set tSum = tSales.GroupBy("Dept", "Amt:SUM, ID:COUNT").OrderBy("Amt_SUM", True)
tSum.DumpToSheet Sheet3.Range("A1"), True
```

`GroupBy`만 새 테이블을 반환하고, 나머지 변형 메서드(`Compute`/`OrderBy`/`AddColumn` 등)는
자기 자신을 반환해 체이닝됩니다(원본이 직접 바뀝니다).

---

## 5. 조회하기 (Query)

### 필터 쌓기

```vb
Dim tRes As clsTable
Set tRes = UnivDB.Query("Emp") _
                 .Where("Dept", "Sales") _
                 .Filter("Salary", ">=", 5000) _
                 .Top(10) _
                 .ToTable()
```

- `.Where(컬럼, 값)` - 해시 인덱스 기반 동등검색. **가장 빠르므로 맨 먼저** 호출해
  후보를 크게 줄이는 것이 좋습니다.
- `.Filter(컬럼, 연산자, 값)` - 조건 스캔. 허용 연산자:
  `=` `<>` `>` `<` `>=` `<=` `IS` `IS NOT` `LIKE` `NOT LIKE`
  (숫자 비교는 `>` 계열, 문자 비교는 `IS`/`LIKE` 계열)
- `.OrderBy(컬럼, [내림차순])` / `.Top(개수)` / `.SelectColumns("A, B")`
- 체인의 조건들은 전부 **AND**로 결합됩니다.

### OR 조건과 IN (FilterAny)

`.FilterAny`는 한 호출이 하나의 OR 그룹입니다. 각 조건(arm)은 `Array(컬럼, 연산자, 값)`.

```vb
' (Dept = "Sales") OR (Amt > 1000) - 그룹 밖 다른 조건과는 AND
Set tRes = UnivDB.Query("Emp") _
                 .FilterAny(Array("Dept", "IS", "Sales"), Array("Amt", ">", 1000)) _
                 .ToTable()

' 같은 컬럼을 나열하면 SQL IN: Dept IN ("Sales", "IT")
Set tRes = UnivDB.Query("Emp") _
                 .FilterAny(Array("Dept", "IS", "Sales"), Array("Dept", "IS", "IT")) _
                 .ToTable()
```

### 쓰기 (Insert / Update / Delete)

쓰기 종결자는 처리한 행 수(Long)를 반환합니다. 값 설정은 `.SetVal`.

```vb
Dim n As Long

' INSERT: 새 행 추가
n = UnivDB.Query("Emp") _
          .SetVal("ID", 101).SetVal("Name", "Dave").SetVal("Dept", "IT") _
          .Insert()

' UPDATE: 조건 매칭 행 일괄 수정
n = UnivDB.Query("Emp") _
          .Where("Name", "Bob") _
          .SetVal("Dept", "Finance") _
          .Update()

' DELETE: 조건 매칭 행 삭제
n = UnivDB.Query("Emp").Filter("Status", "IS", "Retired").Delete()
```

---

## 6. 보고서 만들기 (Report)

`Report`는 원본을 건드리지 않고(Copy-on-Write) 단계별로 **즉시** 가공한 뒤 시트에
서식까지 입혀 출력하는 ETL 빌더입니다.

```vb
UnivDB.Report("SalesData") _
      .Exclude("Status", "IS", "Cancelled") _
      .Keep("Year", "=", 2024) _
      .AddColumns("Grade") _
      .PatchIf("Amount", ">=", 10000, "Grade", "VIP") _
      .SelectColumns("Date", "Customer", "Amount", "Grade") _
      .DefineColumn("Date", vWidth:=12, vFormat:="yyyy-mm-dd", vAlign:="Center") _
      .DefineColumn("Amount", vWidth:=15, vFormat:="#,##0", vAlign:="Right") _
      .BuildToSheet ThisWorkbook.Sheets("Report"), "A1"
```

주요 메서드:

| 메서드 | 역할 |
|---|---|
| `.Keep(컬럼, 연산자, 값)` / `.Exclude(...)` | 조건 매칭 행 유지 / 제거 (즉시 실행) |
| `.KeepAny(Array(...), ...)` | OR 그룹 행 유지 (Query의 FilterAny와 같은 규칙) |
| `.PatchIf(조건컬럼, 연산자, 조건값, 대상컬럼, 새값)` | 조건 매칭 행의 값 일괄 변경 |
| `.AddColumns("A, B")` / `.RemoveColumns("A, B")` | 컬럼 추가 / 삭제 (쉼표 구분 문자열) |
| `.JoinTable(참조테이블, 키, [JoinType], [접두사])` | 참조 테이블 결합 (**기본 jtLeft**) |
| `.FillFromMap(대상컬럼, 소스컬럼, [기본값])` | 소스 컬럼 값으로 채움, 빈 값은 기본값 (VLOOKUP 효과) |
| `.ConvertDates("Col1", "Col2", ...)` | 텍스트(YYYYMMDD 등) → 날짜 변환 |
| `.SelectColumns(...)` → `.DefineColumn(...)` → `.BuildToSheet 시트, "A1"` | 최종 컬럼·서식·출력 |
| `.ToTable()` | 시트 출력 대신 가공 결과를 clsTable로 받기 |

`.JoinTable`의 첫 인자는 clsTable 객체입니다. 등록된 테이블은 `UnivDB.GetTable("이름")`으로
꺼내 넘기세요:

```vb
Set tRes = UnivDB.Report("SalesData") _
                 .JoinTable(UnivDB.GetTable("Depts"), "Dept", , "D_") _
                 .ToTable()
```

---

## 7. 테이블 합치기 (Join)

```vb
' UnivDB.Join(왼쪽, 오른쪽, 왼쪽키, [JoinType], [접두사], [오른쪽키]) -> clsTable
Set tResult = UnivDB.Join("Employees", "Departments", "Dept", jtLeft, "D_")
```

- 왼쪽/오른쪽 자리에 테이블명과 clsTable 객체 혼용 가능
- 접두사(`"D_"`)를 주면 오른쪽 컬럼이 `D_Name`처럼 붙어 이름 충돌을 막습니다
- 오른쪽 키 컬럼명이 다르면 마지막 인자로 지정 (생략 시 왼쪽 키와 동일)

### JoinType 8종

| Enum | 의미 |
|---|---|
| `jtInner` | 교집합 (양쪽 모두 있는 행) |
| `jtLeft` | 왼쪽 전체 + 매칭 (미매칭은 빈칸) - 보고서에서 가장 흔함 |
| `jtRight` | 오른쪽 전체 + 매칭 |
| `jtFull` | 합집합 |
| `jtLeftAnti` | 왼쪽에만 있는 행 (오른쪽에 없는 것) |
| `jtRightAnti` | 오른쪽에만 있는 행 |
| `jtLeftSemi` | 오른쪽에 존재하는 왼쪽 행 (오른쪽 컬럼은 안 붙음) |
| `jtRightSemi` | 왼쪽에 존재하는 오른쪽 행 |

> **기본값 주의**: `UnivDB.Join`과 `Query(...).Join`의 기본은 **jtInner**,
> `Report(...).JoinTable`의 기본은 **jtLeft**입니다.

### Cascade Join 패턴 (참조 테이블 연쇄 확장)

```vb
Set tMain = UnivDB.Join(tMain, "tbl_Project", "Ref_ProjectID", jtLeft, "Proj_")
Set tMain = UnivDB.Join(tMain, "tbl_Client",  "Ref_ClientID",  jtLeft, "Client_")
' -> tMain에 Proj_Name, Client_Name 등 접두사 컬럼이 추가됨
```

---

## 8. 행 조립과 병합 (Builder + Upsert)

행을 하나씩 만들어(또는 고쳐서) 기존 테이블에 병합하는 흐름입니다.

```vb
' 기존 테이블 스키마를 복제한 빌더 (키 컬럼 지정)
Dim B As clsTableBuilder
Set B = UnivDB.EditTable("Inventory", "PROD_ID")

' 커서(Row)를 옮겨가며 값 설정 - 없는 ID는 새 행, 없는 컬럼은 자동 추가
B.Row("P001").SetVal "Qty", 50
Call B.Row("P003").SetVal("Name", "Monitor").SetVal("Qty", 5).SetVal("Location", "A-1")

Debug.Print B.GetVal("Location")   ' 커서 행 값 읽기 (미지 컬럼이면 즉시 오류)

' 조립 결과를 원본에 키 기준 병합 (있으면 Update, 없으면 Insert)
UnivDB.Upsert "Inventory", B.ToTable(), "PROD_ID"
```

- 빈 스키마에서 시작하려면 `UnivDB.Builder("키컬럼")`.
- 원본의 특정 행을 복사해 고치려면 `B.ExtractRow(원본테이블, "ID")`.
- 병합은 파사드 `UnivDB.Upsert 대상, 소스, 키` 또는 `대상테이블.UpsertFrom 소스, 키` -
  같은 동작입니다.

---

## 9. 자주 걸리는 함정

**1) 체인 마지막 호출의 괄호 - VBA 구문 제약**

문(statement) 위치에서 괄호 붙은 다인자 호출로 체인을 끝내면 구문 오류입니다.
(1) 마지막 호출만 괄호 없이 쓰거나 (2) 전체를 `Call`로 감싸세요:

```vb
B.Row("P001").SetVal "Qty", 50                        ' O - 마지막만 괄호 생략
Call B.Row("P001").SetVal("Name", "A").SetVal("Qty", 1)  ' O - Call로 감싸기
' B.Row("P001").SetVal("Qty", 50)                     ' X - 구문 오류
```

**2) Query는 지연, Report는 즉시**

`Query`의 필터는 종결자(`ToTable`/`Update` 등)를 부를 때까지 실행되지 않습니다.
반면 `Report`의 `Keep`/`Exclude` 등은 호출 즉시 작업 테이블이 바뀝니다.

**3) 빈 셀과 오류 셀의 취급**

- 빈 셀(Empty)은 NULL로 취급됩니다: 숫자 비교(`>` 등)와 집계(SUM/AVG 등)에 참여하지 않음
- 오류 셀(#N/A 등)은 어디서도 값으로 참여하지 않고 건너뜁니다(skip)
- 상세 규칙은 `UnivDB_Reference.md` 6장(Value & Key Semantics) 참조

**4) 오타는 즉시 오류 (fail-fast)**

컬럼명 오타, 허용 목록 밖 연산자는 조용한 0행이 아니라 즉시 `Err 9999`를 냅니다.
매크로가 멈추면 메시지의 컬럼명/연산자를 먼저 확인하세요.

**5) InitDB를 썼으면 반드시 ClearDB**

`InitDB`를 호출한 프로시저는 3장 패턴처럼 `Finally` 블록에서 `UnivDB.ClearDB`로
메모리를 정리하고 끝나야 합니다. 프로시저 밖에 clsTable을 살려두지 마세요.

---

## 10. 더 알아보기

- **`UnivDB_Reference.md`** - 전체 API 계약 명세 (연산자 허용 목록, NULL/키 의미론,
  어휘 지도, 표준 코드 패턴). 이 101에서 다루지 않은 세부 규칙은 전부 여기에 있습니다.
