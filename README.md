# UnivDB

**Excel 데이터를 메모리에서 조회·조인·집계하고, VBA 메서드 체인으로 보고서까지.**

[![Release](https://img.shields.io/github/v/release/rfastball/UnivDB?color=217346)](https://github.com/rfastball/UnivDB/releases/latest)
![Excel VBA](https://img.shields.io/badge/Excel-VBA-217346)
[![MIT License](https://img.shields.io/badge/license-MIT-555555)](LICENSE)

[배포본 다운로드](https://github.com/rfastball/UnivDB/releases/latest) · [사용자 가이드](UnivDB_101.md) · [API 명세](UnivDB_Reference.md)

UnivDB는 Excel 배치 작업의 **적재 → 변환 → 보고**를 위한 인메모리 관계 대수 라이브러리입니다. 시트의 표를 배열로 읽고, 조건 검색·조인·집계를 거쳐 다시 시트로 출력합니다. 1만~10만 행 규모의 배치 워크로드를 설계 대상으로 삼습니다.

```vb
UnivDB.Query("Emp") _
      .Where("Dept", "Sales") _
      .Filter("Salary", ">=", 5000) _
      .OrderBy("Salary", True) _
      .SelectColumns("Name", "Salary") _
      .ToTable() _
      .DumpToSheet Sheet2.Range("A1"), True
```

영업 부서에서 급여가 5,000 이상인 직원을 찾아, 급여 내림차순으로 이름과 급여를 출력합니다. 셀마다 반복문을 쓰는 대신 **무엇을 조회하고 어떤 형태로 출력할지**를 표현합니다.

## 시작하기

### 1. 라이브러리 참조 추가

1. [최신 릴리스](https://github.com/rfastball/UnivDB/releases/latest)에서 ZIP을 내려받아 압축을 풉니다. 문서와 라이선스도 함께 보관하세요.
2. `UnivDB.xlam`을 고정 경로에 둡니다.
3. 사용할 매크로 워크북에서 **Alt+F11 → 도구 → 참조 → 찾아보기**로 `UnivDB.xlam`을 선택합니다.

외부 참조에서 빈 테이블을 만들 때는 `New clsTable` 대신 팩토리를 사용합니다.

```vb
Dim t As clsTable
Set t = UnivDB.CreateTable()
```

### 2. 입력 표 준비

매크로가 들어 있는 워크북에 `Data`, `Result` 시트를 만듭니다. `Data!A1:D5`에 아래 표를 입력하고, `Result`는 예제 출력용 빈 시트로 둡니다. `Salary`는 숫자로 입력하세요.

| Name | Dept | Role | Salary |
|---|---|---|---:|
| Alice | Sales | Manager | 7000 |
| Bob | IT | Engineer | 6000 |
| Carol | Sales | Analyst | 5500 |
| Dave | Sales | Associate | 4500 |

### 3. 매크로 실행

표준 모듈에 다음 코드를 넣고 `QuickStart`를 실행합니다.

```vb
Public Sub QuickStart()
    On Error GoTo ErrHandler

    UnivDB.InitDB
    UnivDB.RegisterTable "Emp", ThisWorkbook.Worksheets("Data").Range("A1:D5")

    UnivDB.Query("Emp") _
          .Where("Dept", "Sales") _
          .Filter("Salary", ">=", 5000) _
          .OrderBy("Salary", True) _
          .SelectColumns("Name", "Salary") _
          .ToTable() _
          .DumpToSheet ThisWorkbook.Worksheets("Result").Range("A1"), True

Finally:
    UnivDB.ClearDB
    Exit Sub
ErrHandler:
    MsgBox "오류 " & Err.Number & ": " & Err.Description
    Resume Finally
End Sub
```

`Result!A1`부터 다음 결과가 출력됩니다.

| Name | Salary |
|---|---:|
| Alice | 7000 |
| Carol | 5500 |

사용 흐름은 **`InitDB` → `RegisterTable` → 작업 → 출력 → `ClearDB`**입니다. 오류가 나더라도 마지막 정리 단계에 도달하도록 구성합니다.

> **입출력 범위:** 시작 셀 하나만 등록하면 시트의 마지막 사용 행·열까지 자동 탐색합니다. 여러 표가 있는 시트에서는 예제처럼 헤더를 포함한 범위를 명시하세요. 호출부에서 `.CurrentRegion`을 붙이면 빈 행·열에서 잘린 범위가 그대로 전달될 수 있습니다. `DumpToSheet`는 기본적으로 출력 시작 셀의 기존 `CurrentRegion` 내용을 지우므로, 인접 수식이나 보조열이 없는 전용 출력 영역을 사용하세요. 보존이 필요한 경우의 `bClearRegion` 옵션은 [API 명세](UnivDB_Reference.md)를 참고하세요.

## 작업에 맞는 API

- **조회·CRUD — `Query`**: 조건을 쌓고 `ToTable`, `Insert`, `Update`, `Delete`에서 실행합니다.
- **보고서 — `Report`**: 단계별로 즉시 가공하고 서식을 입혀 출력합니다. Copy-on-Write로 원본을 보호합니다.
- **행 조립 — `Builder`, `EditTable`**: 행과 값을 조립한 뒤 `Upsert`로 키 기준 병합합니다.
- **테이블 결합 — `Join`**: Inner·Left·Right·Full·양방향 Anti/Semi의 8종 조인을 제공합니다.
- **집계·컬럼 연산 — `clsTable`**: `GroupBy`는 새 테이블을 반환하고, `Compute`·`OrderBy` 등은 원본을 변경합니다.

`Query`와 `Report`에는 등록한 테이블 이름 또는 `clsTable` 객체를 넘길 수 있습니다. 조회를 넘어선 사용법과 완결 예제는 [사용자 가이드](UnivDB_101.md)에 있습니다.

## 결과를 신뢰하기 위한 계약

컬럼명 오타나 잘못된 연산자를 조용한 빈 결과로 숨기지 않습니다. 코드의 실수는 즉시 오류로 알리고, 데이터의 빈 값과 오류 값은 정해진 규칙으로 다룹니다.

| 항목 | 동작 |
|---|---|
| 잘못된 컬럼명·연산자·집계 함수 | `Err 9999`로 즉시 실패 |
| 빈 셀 `Empty` | NULL로 취급. 숫자 비교·집계에서 제외하며, 인덱스 키와 문자 비교 경로에서는 빈 문자열로 취급 |
| 오류 셀 `Error` | 비교·필터·집계에서 제외. 정렬에서는 뒤로 배치 |
| 인덱스 키 | 숫자 `42`와 문자 `"42"`를 같은 키로 취급하고, 대소문자를 구분하지 않음 |
| 날짜 | Date Variant 키는 ISO 형식으로 정규화. 시트는 `Value2`로 읽으므로 날짜 서식 셀도 숫자 일련값으로 적재 |
| 평균 `AVG` | 합계에 실제로 가산된 숫자 셀 수를 분모로 사용 |
| 오래된 행 번호·인덱스 | 테이블 세대가 달라지면 `Err 9998`로 실패 |

문자 조건은 `IS`·`LIKE` 계열, 숫자 조건은 `=`·`>=` 등의 비교 연산자를 사용합니다. 세부 규칙은 [API 명세의 Value & Key Semantics](UnivDB_Reference.md#6-value--key-semantics-확정-계약)를 확인하세요.

## 내부 동작

**공개 API → rowID 기반 실행 → 해시·배열 커널**의 3층 구조입니다.

- **테이블:** `clsTable`의 행 중심 2차원 Variant 배열에 데이터를 보관합니다.
- **검색·조인:** `Where`는 해시 인덱스로 후보 행을 찾습니다. 내부에서는 rowID를 전달하고, 결과 테이블이 필요한 경계에서 데이터를 구성합니다.
- **유효성:** 행 번호는 특정 테이블의 특정 세대에서만 유효합니다. 구조가 바뀌면 세대를 갱신하고, 이전 인덱스·행 집합·컬럼 리더의 사용을 차단합니다.

기본 인덱스는 `DictHashIndex`입니다. 실험용 `CoreHashIndex`는 명시적으로 선택할 때만 사용하며, 자동 전환하지 않습니다. 성능은 데이터 분포와 작업 구성에 따라 달라집니다.

<details>
<summary>소스 구조와 개발 루프</summary>

| 소스 | 역할 |
|---|---|
| `src/UnivDB.bas` | 공개 파사드와 테이블 등록 |
| `src/clsTable.cls` | 테이블 저장·변형·입출력 |
| `src/clsQuery.cls` | 조회·CRUD 체인 |
| `src/clsReportBuilder.cls` | ETL·보고서 출력 |
| `src/clsTableBuilder.cls` | 행 조립 |
| `src/DbHashIndex.cls` | 인덱스 인터페이스 |
| `src/modEngine.bas` | 조인·집계 실행 |
| `src/modIndexKey.bas` | 키 정규화 |
| `src/Core_*` | 해시·배열 저수준 커널 |
| `src/Test*.bas` | 테스트 스위트 |

소스 정본은 `src/`이며, `UnivDB.xlsm`은 개발 워크북, `UnivDB.xlam`은 배포 라이브러리입니다. Core 커널도 이 저장소에서 관리합니다.

개발 환경에서는 `vbaloop`로 소스를 주입하고 테스트합니다. `Public Sub Test_*()`가 전체 자동 수집됩니다.

```powershell
vbaloop init --check
vbaloop push
vbaloop test
```

환경 준비와 실패 복구 절차는 [개발 규칙](AGENTS.md), 구조와 세대 계약의 근거는 [설계 결정](src/docs/UnivDB_design_decisions_202607.md)을 참고하세요.

</details>

## 더 알아보기

- **[사용자 가이드](UnivDB_101.md)** — CRUD, 보고서, 조인, 행 조립을 예제로 익히기
- **[API 명세](UnivDB_Reference.md)** — 메서드 시그니처, 값·키 계약, 호환 별칭 확인하기
- **[릴리스](https://github.com/rfastball/UnivDB/releases/latest)** — 배포본과 버전별 변경 사항 확인하기

## 라이선스

[MIT License](LICENSE) © 2026 JM. 상업적 이용을 포함해 사용·수정·재배포할 수 있습니다. 재배포 시 저작권 표시와 라이선스 전문을 포함해야 하며, 소프트웨어는 보증 없이 제공됩니다.
