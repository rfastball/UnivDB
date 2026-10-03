# UnivDB

> Excel VBA를 위한 인메모리 관계 대수 엔진 — 테이블 위에서 조회·조인·집계·보고를 메서드 체인으로.

![version](https://img.shields.io/badge/version-v2.8.0-blue) ![platform](https://img.shields.io/badge/platform-Excel%20VBA-217346) ![language](https://img.shields.io/badge/language-VBA-lightgrey)

UnivDB는 Excel 배치 워크로드(1만~10만 행 규모의 **적재 → 변환 → 보고**)를 위한 인메모리 관계 대수 엔진입니다. 시트 데이터를 메모리로 올려 선택·투영·조인·집합연산·집계한 뒤 시트로 되돌립니다. SQL이 아니라 **메서드 체인**으로 쓰고, 셀 단위 루프 대신 배열 기반으로 동작해 수만 행에서도 빠릅니다.

지향점은 배포용 라이브러리입니다: **API 안정성, 명시적 계약, fail-fast 우선.** "조용히 틀린 결과"를 최악의 실패로 보고, 컬럼명 오타나 잘못된 연산자 같은 코드의 실수는 조용히 0행을 돌려주지 않고 즉시 오류(`Err 9999`)를 냅니다.

---

## 왜 UnivDB인가

일반적인 VBA 데이터 가공은 셀 루프·`Dictionary`·`VLOOKUP` 수식이 뒤섞여 느리고 깨지기 쉽습니다. UnivDB는 그 반복 작업을 하나의 표현력 있는 체인으로 대체합니다.

```vb
UnivDB.Query("Emp") _
      .Where("Dept", "Sales") _
      .Filter("Salary", ">=", 5000) _
      .OrderBy("Salary", True) _
      .SelectColumns("Name, Role, Salary") _
      .ToTable() _
      .DumpToSheet Sheet2.Range("A1"), True
```

- **빠름** — row-major 2D Variant 스토리지 + 해시 인덱스. `Where`는 인덱스 동등검색, 조인은 해시 조인.
- **안전함** — fail-fast 계약. 미지 연산자·컬럼 오타는 즉시 raise, 데이터 셀 수준 문제(Empty/Error)는 명시적 NULL/skip 규칙대로 처리.
- **선언적** — 무엇을 원하는지 쓰면, 어떻게 도는지는 엔진이 결정.

---

## 5분 시작

```vb
Sub QuickStart()
    On Error GoTo ErrHandler

    ' (1) 엔진 초기화 + 시트 데이터 등록 (시작 셀만 주면 범위 자동 인식)
    UnivDB.InitDB
    UnivDB.RegisterTable "Emp", Sheet1.Range("A1")

    ' (2) 조회: Where(인덱스) → Filter(스캔) → 정렬 → 컬럼 선택 → 출력
    UnivDB.Query("Emp") _
          .Where("Dept", "Sales") _
          .Filter("Salary", ">=", 5000) _
          .OrderBy("Salary", True) _
          .SelectColumns("Name, Role, Salary") _
          .ToTable() _
          .DumpToSheet Sheet2.Range("A1"), True

Finally:
    ' (3) 반드시 정리 — InitDB를 호출한 프로시저는 ClearDB로 끝난다
    UnivDB.ClearDB
    Exit Sub
ErrHandler:
    MsgBox "오류 " & Err.Number & ": " & Err.Description
    Resume Finally
End Sub
```

한 매크로에 사용의 뼈대가 다 들어 있습니다: **InitDB → RegisterTable → 작업 → 종결자 → ClearDB.**

---

## 핵심 개념: 테이블 하나, 작업 표면 넷

모든 데이터는 `clsTable`에 담기고, 이를 다루는 진입점이 넷입니다.

| 진입점 | 반환 | 성격 |
|---|---|---|
| `UnivDB.Query(T)` | `clsQuery` | 검색·CRUD (지연 실행, `ToTable()`에서 물질화) |
| `UnivDB.Report(T)` | `clsReportBuilder` | ETL 파이프라인 (Copy-on-Write, 원본 보존) |
| `UnivDB.Builder(Key)` | `clsTableBuilder` | 행 단위 동적 조립 |
| `UnivDB.Join(...)` | `clsTable` | 8종 조인 (Inner/Left/Right/Full/Anti/Semi) |

### Query — 검색과 CRUD

```vb
' UPDATE: 조건 매칭 행 일괄 수정
UnivDB.Query("Emp") _
      .Where("Name", "Bob") _
      .SetValue("Dept", "Finance") _
      .Update
```

교차 컬럼 OR은 `FilterAny`로: `(arm1 OR arm2) AND ...` 형태의 AND-of-ORs를 표현합니다.

### Report — ETL 파이프라인

```vb
UnivDB.Report("SalesData") _
      .Exclude("Status", "=", "Cancelled") _
      .Keep("Year", "=", 2024) _
      .PatchIf("Amount", ">=", 10000, "Grade", "VIP") _
      .SelectColumns("Date", "Customer", "Amount", "Grade") _
      .DefineColumn("Date", vWidth:=12, vFormat:="yyyy-mm-dd", vAlign:="Center") _
      .DefineColumn("Amount", vWidth:=15, vFormat:="#,##0", vAlign:="Right") _
      .BuildToSheet Sheet2, "A1"
```

### Join — 캐스케이드 조인

```vb
Set tMain = UnivDB.Join(tMain, "tbl_Project", "Ref_ProjectID", jtLeft, "Proj_")
Set tMain = UnivDB.Join(tMain, "tbl_Client",  "Ref_ClientID",  jtLeft, "Client_")
' → tMain.headers에 "Proj_Name", "Client_Name" 등 접두사 컬럼 추가
```

### 집계 (GroupBy)

```vb
Set tTop = T.GroupBy("Dept", "Amt:SUM, Headcount:COUNT").OrderBy("Amt_SUM", True)
```

---

## 아키텍처

3층 구조입니다.

1. **공개 파사드** — `UnivDB.bas` (InitDB / RegisterTable / Query / Report / Join / Builder)와 메서드 체이닝 객체들. 정본 스토리지는 `clsTable`의 row-major 2D Variant.
2. **내부 실행 언어** — 쿼리 실행은 rowID 중심으로 흐르고 파사드 경계에서만 물질화. 인덱스 진실은 `DbHashIndex` 인터페이스 하나(기본 `DictHashIndex`, 실험 백엔드 `CoreHashIndex`).
3. **Core 커널** — `Core_*` 해시/배열 저수준 프리미티브. DB 의미론 없음.

### 세대(Generation) 계약 — 최상위 불변식

rowID는 **특정 테이블의 특정 세대 안에서만 유효한 물리 행 번호**입니다(surrogate ID 아님). 모든 구조 변경(Insert/Delete/OrderBy/Upsert/컬럼 추가·삭제)은 `Generation`을 증가시키고, 모든 파생 구조(인덱스·RowSet·RowList·ColumnReader)는 생성 시점의 (테이블, 세대)에 결속됩니다. 불일치 시 `Err.Raise 9998`로 fail-fast — 세대가 어긋난 인덱스로 "조용히 틀린 결과"를 내는 일을 구조적으로 차단합니다.

---

## 값·키 의미론 (fail-fast 계약)

"조용히 틀린 결과"를 막기 위한 값/키 계약이 인덱스·필터·정렬·집계 전 표면에서 동일하게 적용됩니다.

- **Empty = NULL** — 숫자 비교·모든 집계 불참. 문자 경로(IS/LIKE)와 인덱스 키에서만 `""`로 참여.
- **Error = skip** — 비교/정렬/집계/필터 어디서도 값으로 참여하지 않음. OrderBy는 크래시 없이 맨 뒤 배치.
- **인덱스 키 정규화** — 숫자 `42`와 문자 `"42"`는 같은 키로 병합, Date Variant는 ISO canonical(`yyyy-mm-dd`, locale 무관). `LoadFromSheet`는 `Value2`를 사용하므로 날짜 서식 셀도 숫자 일련값으로 적재합니다. 키 비교는 대소문자를 무시합니다.
- **연산자 화이트리스트** — 미지 연산자/집계함수 문자열은 즉시 `Err.Raise 9999`. 조용한 0행 없음.

`AVG = SUM / (숫자로 가산된 셀 수)` — 비숫자 실값이 섞여도 평균이 희석되지 않습니다.

---

## 설치

배포 라이브러리(`UnivDB.xlam`)를 참조로 추가해 사용합니다.

1. `UnivDB.xlam`을 고정 경로에 둡니다.
2. 내 워크북 VBE(Alt+F11) → **도구 > 참조 > 찾아보기**로 선택해 체크합니다.
3. 이제 `UnivDB.InitDB`, `Dim t As clsTable` 등을 바로 씁니다.

> **중요**: 참조 방식에서는 `New clsTable`을 직접 쓸 수 없습니다. 빈 테이블은 팩토리로:
> ```vb
> Dim t As clsTable
> Set t = UnivDB.CreateTable()   ' O
> ' Dim t As New clsTable        ' X — 참조 프로젝트 밖에서는 컴파일 오류
> ```

---

## 문서

| 문서 | 내용 |
|---|---|
| [UnivDB_101.md](UnivDB_101.md) | 처음 시작하는 사용자 가이드 (설치·5분 시작·핵심 개념) |
| [UnivDB_Reference.md](UnivDB_Reference.md) | 전체 API 계약 명세 (파사드·4개 표면·값/키 의미론·조인·템플릿) |

---

## 프로젝트 구조

```
src/
  UnivDB.bas            공개 파사드 진입점
  clsTable.cls          정본 스토리지 (row-major 2D Variant)
  clsQuery.cls          검색·CRUD 체인
  clsReportBuilder.cls  ETL 파이프라인
  clsTableBuilder.cls   행 단위 동적 조립
  DbHashIndex.cls       인덱스 인터페이스 (DictHashIndex / CoreHashIndex)
  modEngine.bas         조인·집계 실행 엔진
  modIndexKey.bas       키 정규화 단일 계약
  modOperators.bas      연산자 화이트리스트 (fail-fast)
  Core_*                해시/배열 저수준 프리미티브 (DB 의미론 없음)
  Test*.bas             회귀 테스트 스위트
```

---

## 상태

버전 **v2.8.0**. 공개 API는 안정적이며 레거시 이름은 영구 유지됩니다(어휘 지도는 Reference 6.7절 참조). 소스는 `src/*.bas`, `src/*.cls`가 진실이고, `UnivDB.xlsm`은 사람이 관리하는 런타임 아티팩트입니다.
