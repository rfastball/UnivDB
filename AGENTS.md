# AGENTS.md / CLAUDE.md - UnivDB

(시그니처 스키마·src 편집 주의는 상위 `VBA_Projects/CLAUDE.md` 참조. 이 파일과 AGENTS.md는 동일 내용으로 유지)

대상 워크북: UnivDB.xlsm (구 dev_UnivDB.xlsb에서 이름 변경 및 xlsm 마이그레이션됨)

테스트는 `Public Sub Test_*()`가 전량 자동 수집·실행된다(개별 선택 실행 없음).
벤치마크는 `TestBenchDb.bas` - 결과는 `LogShim.LogJ`로 jsonl에 baseline 보존.

## 의존성: UnivCore (deps.lock으로 버전 고정)

- src\Core_* 는 설치 산출물 (gitignore, 직접 수정 금지 - 수정은 UnivCore에서)
- `vh deps check/install/update` 자동 hydrate는 하네스 기능이라 현재 미사용 — 당분간
  UnivCore → UnivDB Core_* 동기화는 수동으로 파일 복사해서 처리

## 무엇을 만드는가

Excel 배치 워크로드(1만~10만 행, 적재→변환→보고)용 인메모리 DB 라이브러리.
지향점은 배포용 라이브러리: API 안정성, 명시적 계약, fail-fast 우선.
"조용히 틀린 결과"가 최악의 실패 양식이다.

## 아키텍처 3층

1. **공개 파사드** - `UnivDB.bas` (InitDB / RegisterTable / Query / Report / Join / Builder)
   + 메서드 체이닝 객체 `clsQuery`(CRUD), `clsReportBuilder`(ETL), `clsTableBuilder`(행 조립).
   정본 스토리지는 `clsTable`의 row-major 2D Variant (`m_vData`). 공개 API 문서: `UnivDB_Reference.md`
2. **내부 실행 언어** - 쿼리 실행은 rowID 중심으로 흐르고 파사드 경계에서만 물질화한다.
   - `DbHashIndex.cls` = 백엔드 교체용 인터페이스. 기본 구현 `DictHashIndex`,
     실험 백엔드 `CoreHashIndex`(FastHash 기반, `clsTable.IndexBackend = dbBackendCore` 명시 옵션 전용)
   - 보조 구조: `DbRowList` / `DbRowSet` / `DbIndexCursor` / `DbColumnReader`(컬럼 접근 밀봉 표면)
   - 키 정규화는 `modIndexKey.NormalizeIndexKey` 단일 계약: Empty→"", Error 키 skip, CStr, vbTextCompare(대소문자 무시)
3. **Core 커널** - `Core_*` (UnivCore 공급). 해시/배열 저수준 프리미티브만. DB 의미론 없음.

## 세대(Generation) 계약 - 최상위 불변식

rowID는 특정 테이블의 특정 세대 안에서만 유효한 물리 행 번호다 (surrogate ID 아님).

- 모든 구조 변경(Insert/Delete/OrderBy/Upsert/컬럼 추가·삭제)은 `clsTable.Generation`을 증가시킨다
- 모든 파생 구조(HashIndex, RowSet, RowList, ColumnReader)는 생성 시점 (테이블, 세대)에 결속된다
  (`modDbContract`가 검증, 불일치 시 Err.Raise 9998 fail-fast)
- 새 파생 구조를 추가할 때의 질문은 하나다: "세대를 기억하는가, 불일치 시 실패하는가"

## 하류 벤더링 (UnivFilter)

- `src/clsSet.cls`와 `src/modIndexKey.bas`는 **UnivFilter가 스냅샷 벤더링하는 공유 요소**다
  (편집 본가 = 여기 UnivDB, 라이브 싱크 없음, 드리프트 허용, 출처는 UnivFilter의
  `src/vendor.manifest`에 UnivDB HEAD 해시로 기록).
- clsSet은 dictHash 계보의 "집합연산 표면"으로 공식화한다. 단, **쿼리 실행 경로에서의
  clsSet 사용 금지는 그대로 유지**된다 (아래 금지 사항 참조).
- 두 파일의 공개 API/키 의미론(Variant 키: 숫자 1과 문자 "1" 구분, vbTextCompare)을
  바꿀 때는 UnivFilter 쪽 재스냅샷 필요 여부를 함께 판단할 것.

## 금지 사항 (설계 결정으로 확정됨)

- Core에 DB 의미론 주입 금지 (`Core_JoinTables`, `Core_GroupBy` 류)
- 쿼리 실행 경로에 `someDict.Keys` / `New Collection` / `clsSet` 재도입 금지
- raw `Dictionary(key -> Collection)` 인덱스 경로 재도입 금지 (인덱스 진실은 `DbHashIndex` 하나)
- public 호환 shim(`SelectRows` 등) 시그니처 변경 금지
- 백엔드 자동 선택 도입 금지 - 기본은 DictHashIndex (policy v2, 2026-07-07 재확정),
  Core는 build/join/cursor/GroupBy-heavy 명시 옵션 + 벤치 기반 수동 판단

## 알려진 미결 사항

- **Core 백엔드 컴파일 결함 (2026-07-08 확인)**: Core 백엔드 최초 컴파일 그래프
  (CoreHashIndex / Core_clsFastHash / Core_clsLongList / Core_modMemCore 중 하나)에
  구문 오류가 있어, Core 경로를 건드리는 테스트(TestBenchDb, *_Parametric,
  *_CoreBackend_*)가 VBE 컴파일 오류 모달을 띄우고 vbaflow 세션을 죽인다.
  dict 경로 테스트는 전부 green. clsSet과는 무관(가드 보강 전에도 동일 재현).
  수정은 UnivCore 동기화 소관.

- GroupBy 키 의미론은 인덱스 계약과 통일됨 (R1/D2 결정, 2026-07-07). 단일 컬럼은 인덱스 커서
  기반 엔진화 (R2), 복합 키는 길이 접두 인코딩으로 충돌-프리 (R3). 복합 경로의 누적기 엔진화와
  hot column cache(R4 보류)는 병목 관측 시 재평가
- NULL/타입 의미론(비결정 D)은 1.0 전 확정 필요 - GroupBy 그룹 컬럼 값의 String화도 이 소관.
  실행 계획(D0~D3)은 `src/docs/tasks/2026-07-08_null_type_semantics.md` (방향 승인됨)

## 설계 문서 (읽는 순서)

- `src/docs/tasks/2026-07-08_null_type_semantics.md` - 다음 작업 축 (비결정 D 실행 계획 D0~D3)
- `src/docs/tasks/2026-07-05_next_direction.md` - Phase R0~R5 기록 (완료된 이력 + 결정 로그)
- `src/docs/UnivDB_design_decisions_202607.md` - 세대 계약·스토리지 레이아웃 결정과 근거
- `src/docs/UnivDB_refactor_plan.md` - 내부 실행 표현 전환 원 설계 (Phase 0~8, 완료된 이력)

<!-- vbaflow:managed:start -->
## 개발 루프 (_VbaFlow)

빌드/테스트 루프는 `vbaflow`(C# .NET 8 단일 exe)로 수행한다. PATH 미등록 시 exe 경로:
`..\_VbaFlow\src\VbaFlow\bin\Release\net8.0-windows\win-x64\publish\vbaflow.exe`
온보딩/드리프트 복구: `vbaflow init` (검사만: `vbaflow init --check`).
프로젝트 고유 내용은 이 관리 블록 밖에 작성한다 (init이 블록 내부를 덮어씀).

1. 편집: src\*.bas/.cls/.frm 수정
2. 세션: vbaflow session start (도구 소유 Excel, 멱등)
3. 주입: vbaflow push (src -> 워크북 in-place 주입, --prune 지원)
4. 실행: vbaflow test (Test_* 전체 실행 + LogShim jsonl 판독, --timeout <초>)
5. 수정: fail이면 원인 고치고 1로, green이면 git commit

- 모든 명령 --json 지원: {command, ok, exitCode, hint, errors[], payload}. 성공만 exit 0
- 테스트는 Public Sub Test_*() (인자 없음), 단언은 LogShim.AssertEq (Debug.Assert 금지)
- src\LogShim.bas는 벤더링 사본 - 직접 수정 금지, 갱신은 vbaflow init
- pull은 src의 미푸시 편집을 덮어쓴다 - 워크북 쪽이 진실일 때만
- src 규약: UTF-8 BOM 없음 + CRLF + 끝 개행 1개. CP949 밖 문자(em dash U+2014, 이모지)는 push가 거부
<!-- vbaflow:managed:end -->
