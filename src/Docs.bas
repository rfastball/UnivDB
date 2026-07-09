Attribute VB_Name = "Docs"
Option Explicit

' ==============================================================================
' [UnivDB API Reference v2.1]
' ==============================================================================
' 작성자: UnivDB Project Team
' 업데이트: 2024-XX-XX (CRUD, ReportBuilder, DefineColumn 반영)
'
' [목차]
' 1. UnivDB (Facade)      - 라이브러리 진입점
' 2. clsQuery             - 데이터 조회 및 조작 (CRUD)
' 3. clsReportBuilder     - 데이터 가공 파이프라인
' 4. clsTable             - 데이터 컨테이너 및 집계
' 5. 상세 문법 가이드 (Syntax)
' ==============================================================================

' ------------------------------------------------------------------------------
' [1] UnivDB (Facade Module)
' 설명: 라이브러리 초기화 및 주요 기능의 진입점입니다.
' ------------------------------------------------------------------------------
' 1. InitDB()
'    - DB 엔진을 초기화합니다. (최초 1회 필수)
'
' 2. RegisterTable(Name, Source)
'    - 시트 범위(Range)나 배열(Array)을 테이블 이름으로 등록합니다.
'    - 예: UnivDB.RegisterTable "Emp", Sheet1.Range("A1").CurrentRegion
'
' 3. Query(Target) -> clsQuery
'    - 쿼리 빌더를 시작합니다. Target은 테이블명(String) 또는 객체(clsTable)입니다.
'    - 예: UnivDB.Query("Emp").Where(...).SelectTable
'
' 4. Report(Target) -> clsReportBuilder
'    - 리포트 빌더를 시작합니다. (Factory Method)
'    - 예: UnivDB.Report("Emp").Exclude(...).BuildToSheet ...
'
' 5. Join(Left, Right, Key, [Type], [Prefix], [RightKey]) -> clsTable
'    - 두 테이블을 조인하여 새로운 테이블을 반환합니다.
'    - Type: jtInner(1), jtLeft(2), jtRight(3), jtFull(4), jtLeftAnti(5) 등
' ------------------------------------------------------------------------------

' ------------------------------------------------------------------------------
' [2] clsQuery (Query Builder)
' 설명: 데이터 검색, 정렬, 추출 및 CRUD 작업을 수행합니다.
' ------------------------------------------------------------------------------
' [검색 및 필터]
' - .Where(Col, Val)             : 인덱스 기반 동등 검색 (가장 빠름)
' - .Filter(Col, Op, Val)        : 조건 검색 (>, <, <>, LIKE 등)
' - .FilterIn(Col, Val1, ...)    : 포함 검색 (SQL IN)
' - .FilterNotIn(Col, Val1, ...) : 제외 검색 (SQL NOT IN)
' - .FilterLike(Col, Pat1, ...)  : 패턴 포함 (하나라도 매칭되면 선택)
' - .FilterNotLike(Col, Pat...)  : 패턴 제외 (하나라도 매칭되면 제외)
'
' [정렬 및 제한]
' - .OrderBy(Col, [Desc])        : 정렬 (Desc=True 시 내림차순)
' - .Top(Count)                  : 상위 N개 행만 선택
' - .SelectColumns(Cols)         : 출력할 컬럼 지정 (쉼표 구분)
'
' [실행 (Read)]
' - .SelectTable() -> clsTable   : 조회 결과를 테이블 객체로 반환
'
' [실행 (Write / CRUD)]
' - .SetValue(Col, Val)          : Insert/Update용 값 설정 (체이닝 가능)
' - .Insert() -> Long            : 행 추가 (SetValue 값 사용)
' - .Update() -> Long            : 조건에 맞는 행 수정 (SetValue 값 사용)
' - .Delete() -> Long            : 조건에 맞는 행 삭제
' ------------------------------------------------------------------------------

' ------------------------------------------------------------------------------
' [3] clsReportBuilder (ETL Pipeline)
' 설명: 원본 데이터를 가공하여 최종 리포트를 생성합니다. (메서드 체이닝 최적화)
' ------------------------------------------------------------------------------
' [구조 변경]
' - .AddColumns(ColList)         : 빈 컬럼 추가 (쉼표 구분 문자열)
' - .RemoveColumns(ColList)      : 컬럼 삭제 (쉼표 구분 문자열)
' - .CreateKey(New, Src, Len)    : 키 컬럼 생성 (Left 함수 적용)
'
' [데이터 결합 및 변환]
' - .JoinTable(Ref, Key, [Type], [Prefix], [RightKey]) : 다른 테이블과 조인
' - .FillFromMap(Tgt, Src, Def)  : 조인된 데이터로 값 채우기 (VLOOKUP 효과)
' - .ConvertDates(Cols...)       : 텍스트 날짜를 실제 날짜형으로 변환
'
' [필터링 및 값 수정]
' - .Filter(Col, Op, Val)        : 조건에 맞는 행만 유지 (Keep)
' - .Exclude(Col, Op, Val)       : 조건에 맞는 행 삭제 (Remove)
' - .UpdateValue(Tgt, Val, CondCol, Op, CondVal) : 조건부 값 변경
'
' [출력 설정]
' - .SelectFinal(Cols...)        : 최종 출력 컬럼 및 순서 결정
' - .DefineColumn(Col, [Width], [Format], [Align], [Alias]) : 컬럼 서식 정의
' - .BuildToSheet(Sheet, Addr)   : 엑셀 시트에 출력 (서식 적용 포함)
' ------------------------------------------------------------------------------

' ------------------------------------------------------------------------------
' [4] clsTable (Core Container)
' 설명: 데이터를 메모리에 저장하고 관리하는 핵심 객체입니다.
' ------------------------------------------------------------------------------
' [입출력]
' - .LoadFromSheet(Range)        : 시트에서 로드
' - .LoadFromArray(Array)        : 배열에서 로드
' - .DumpToSheet(Range)          : 시트로 출력
'
' [데이터 관리]
' - .UpsertFrom(Table, Key)      : 다른 테이블의 데이터로 갱신 (Merge)
' - .GroupBy(Cols, AggList)      : 그룹별 집계 (SUM, AVG, COUNT, MAX, MIN)
' - .RawArray                    : 내부 2차원 배열 접근 (읽기 전용 권장)
' ------------------------------------------------------------------------------

' ==============================================================================
' [5] 상세 문법 가이드 (Syntax Guide)
' ==============================================================================
'
' ------------------------------------------------------------------------------
' [5-1] 필터링 (Filtering)
' ------------------------------------------------------------------------------
' 1) .Where(Column, Value)
'    - 인덱스를 사용하므로 가장 빠릅니다. (동등 비교 전용)
'    - 예: .Where("Dept", "Sales")
'
' 2) .Filter(Column, Operator, Value)
'    - 연산자: ">", "<", ">=", "<=", "<>", "=", "LIKE", "NOT LIKE"
'    - 예: .Filter("Age", ">=", 30)
'    - 예: .Filter("Name", "LIKE", "Kim*")
'
' ------------------------------------------------------------------------------
' [5-2] 서식 정의 (DefineColumn)
' ------------------------------------------------------------------------------
' - vWidth       : 컬럼 너비 (숫자)
' - vFormat      : 엑셀 표시 형식 (예: "#,##0", "yyyy-mm-dd", "@")
' - vAlign       : 정렬 ("Left", "Center", "Right")
' - vHeaderAlias : 헤더 표시명 변경
'
' - 예: .DefineColumn("Amount", vWidth:=12, vFormat:="#,##0", vAlign:="Right")
'
' ------------------------------------------------------------------------------
' [5-3] 집계 함수 (GroupBy)
' ------------------------------------------------------------------------------
' - 형식: "컬럼명:함수" (쉼표로 구분)
' - 지원 함수: COUNT, SUM, AVG, MAX, MIN
' - 예: .GroupBy("Dept", "Salary:AVG, EmpID:COUNT")
'
' ------------------------------------------------------------------------------
' [5-4] 조인 타입 (JoinType Enum)
' ------------------------------------------------------------------------------
' - jtInner (1)     : 교집합
' - jtLeft (2)      : Left Outer Join
' - jtRight (3)     : Right Outer Join
' - jtFull (4)      : Full Outer Join
' - jtLeftAnti (5)  : 차집합 (왼쪽에만 있는 데이터)
' - jtLeftSemi (7)  : 교집합 (왼쪽 데이터만 반환)
' ==============================================================================
