# AGENTS.md / CLAUDE.md - UnivDB

(시그니처 스키마·src 편집 주의는 상위 `VBA_Projects/CLAUDE.md` 참조. 이 파일과 AGENTS.md는 동일 내용으로 유지)

대상 워크북: UnivDB.xlsm (구 dev_UnivDB.xlsb에서 이름 변경 및 xlsm 마이그레이션됨)

테스트는 `Public Sub Test_*()`가 전량 자동 수집·실행된다(개별 선택 실행 없음).
벤치마크는 `TestBenchDb.bas` - 결과는 `LogShim.LogJ`로 jsonl에 baseline 보존.

## Core 커널 (구 UnivCore, 동결 — UnivDB가 정본)

- `src\Core_*` 는 해시/배열 저수준 프리미티브(구 UnivCore 공급분).
- **UnivCore 프로젝트는 동결**되었다. UnivFilter가 UnivDB 스냅샷 기잔이 되면서 분리 명분이
  사라졌고, 이제 **UnivDB가 Core_*의 정본**이다. 상류 전파 대상 없음 — 수정은 여기서 직접 한다.
- 따라서 Core_* 는 더 이상 설치 산출물이 아니라 UnivDB의 1급 소스로 git 추적한다(gitignore 해제됨).
- DB 의미론 주입 금지(Core에 Join/GroupBy 류 금지)는 그대로 유지(아래 금지 사항).

## 무엇을 만드는가

Excel 배치 워크로드(1만~10만 행, 적재→변환→보고)용 인메모리 관계 대수 엔진(테이블 위 쿼리·조인·집계·보고).
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
     / `DbCsrPayload`(인덱스 rowId를 CSR 평탄 배열로 보관 - 키당 객체 없음, 양 백엔드 공유)
   - 키 정규화는 `modIndexKey.NormalizeIndexKey` 단일 계약: Empty→"", Error 키 skip, CStr, vbTextCompare(대소문자 무시)
3. **Core 커널** - `Core_*` (구 UnivCore, 현재 UnivDB 정본). 해시/배열 저수준 프리미티브만. DB 의미론 없음.

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
- **2026-10-31 (2.10.0 성능 계획 단계 6)**: `modIndexKey.bas`에 `FoldIndexKey`(하이브리드 키 접기, LCMapStringW Declare를
  모듈 안에 둠 - Core_* 비의존) 추가 = **UnivFilter 재스냅샷 필요**. clsSet은 이번 범위 밖(접기 미적용).

## 금지 사항 (설계 결정으로 확정됨)

- Core에 DB 의미론 주입 금지 (`Core_JoinTables`, `Core_GroupBy` 류)
- 쿼리 실행 경로에 `someDict.Keys` / `New Collection` / `clsSet` 재도입 금지
- raw `Dictionary(key -> Collection)` 인덱스 경로 재도입 금지 (인덱스 진실은 `DbHashIndex` 하나)
- public 호환 shim(`SelectRows` 등) 시그니처 변경 금지
- 백엔드 자동 선택 도입 금지 - 기본은 DictHashIndex (policy v2, 2026-07-07 재확정),
  Core는 build/join/cursor/GroupBy-heavy 명시 옵션 + 벤치 기반 수동 판단

## 알려진 미결 사항

- **Core 백엔드 컴파일 결함 (2026-07-08 발견 → 2026-07-10 해결)**: Core 경로 테스트(TestBenchDb,
  *_Parametric, *_CoreBackend_*)가 VBE 컴파일 오류 모달로 세션을 죽이던 결함. 원인은 UnivCore
  구문 오류가 아니라 vbaflow push의 AddFromString이 `_` 컨티뉴에이션 Declare 문을 오파싱해
  모듈 끝에 팬텀 `()` 라인을 주입한 것(Core_modMemCore의 RtlMoveMemory Declare에서 발현).
  최소 재현·근본원인·수정 후보는 `..\_VbaFlow\docs\bugreports\` 2026-07-10 리포트 참조.
  수정: `Core_modMemCore.bas`의 Declare 2개를 한 줄로 폄(컨티뉴에이션 제거). UnivDB가 Core_*의
  정본이므로 이 수정이 정본(상류 전파 대상 없음). 해결 후 Core 경로 테스트 전량 green(sessionDied=False).

- GroupBy 키 의미론은 인덱스 계약과 통일됨 (R1/D2 결정, 2026-07-07). 단일 컬럼은 인덱스 커서
  기반 엔진화 (R2), 복합 키는 길이 접두 인코딩으로 충돌-프리 (R3). 복합 경로의 누적기 엔진화와
  hot column cache(R4 보류)는 병목 관측 시 재평가
- **NULL/타입 의미론 (비결정 D, 2026-07-10 확정 완료)**: D0~D3 전부 종결. 확정 계약은
  Empty=NULL + Error=skip 전 표면(비교/정렬/집계/필터) 통일, 날짜 인덱스 키 ISO canonical,
  OrderBy 날짜 수치 비교, AVG 분모=숫자 셀 수. 공개 계약은 `UnivDB_Reference.md`
  "6. Value & Key Semantics" 장, 실행 이력은 `src/docs/tasks/2026-07-08_null_type_semantics.md`.
  **후속(타 프로젝트 몫): `modIndexKey.bas` 키 의미론 변경으로 UnivFilter 재스냅샷 필요**

## 설계 문서 (읽는 순서)

- `src/docs/tasks/2026-07-10_chaining_api_audit.md` - 체이닝 API 전수조사 + 전면 리팩토링
  (위반 F1~F24 교정 완료, 잔여 이월은 7.4절)
- `src/docs/reference/query_api_language_survey_202607.md` - 상용 쿼리 API 언어철학 조사
  + 문법 검증 기준 V1~V9 (위 audit의 근거)
- `src/docs/tasks/2026-07-08_null_type_semantics.md` - 비결정 D 실행 계획 D0~D3 (완료된 이력)
- `src/docs/tasks/2026-07-05_next_direction.md` - Phase R0~R5 기록 (완료된 이력 + 결정 로그)
- `src/docs/UnivDB_design_decisions_202607.md` - 세대 계약·스토리지 레이아웃 결정과 근거
- `src/docs/UnivDB_refactor_plan.md` - 내부 실행 표현 전환 원 설계 (Phase 0~8, 완료된 이력)

<!-- vbaloop:managed:start -->
## Development loop (vbaloop)

The build/test loop runs through `vbaloop` (a single exe). If it is not on PATH, the install path is
`%LOCALAPPDATA%\vbaloop\bin\vbaloop.exe` (if even that is missing, reinstall with `vbaloop setup`).
Onboarding/drift repair: `vbaloop init` (check only: `vbaloop init --check`).
If init --check exits 3 with `exe-older` (the managed_by version in vbaloop.toml is newer than the running exe), do not fix it
with init; ask the user to run `vbaloop update` (do not edit or delete the managed_by line).
Write project-specific content outside this managed block (init overwrites the inside of the block).

1. Edit: change src\*.bas/.cls/.frm
   New UserForm: write JSON and plain VBA outside src, then
   `vbaloop form build --spec forms/MyForm.json --code forms/MyForm.vba --json` -> push.
   No template or hand-written FRX is needed. Before form work, read
   `.claude/skills/vbaloop/references/forms.md` for design, code rules and schema.
   Rebuild with --force only to replace the generated pair from those inputs.
2. Session: ephemeral by default - push/pull/test/run/diff start a hidden Excel per command and close it (vbaloop.toml
   [excel] session). If a session opened with session start/attach is alive, it is used
3. Inject: vbaloop push (src -> in-place injection into the workbook -> saved only if the VBE compile passes; supports --prune/--force.
   Backup to _dbg\backups\ right before injecting)
4. Run: vbaloop test (runs every Test_* + reads the LogShim jsonl, --timeout <seconds>)
5. Fix: on fail, fix the cause and go back to 1; when green you are done (commit now if you use version control)

- Every command supports --json: {command, ok, exitCode, hint, errors[], payload}
- exit code: success 0 / environment error 3 (session/Excel/COM/busy/quarantined - follow the hint) / otherwise 2
- Quarantined (hint mentions recovery clear / --discard): the tool could not safely clean up Excel, so it stopped instead of killing it.
  Never kill Excel on your own initiative - that includes running --discard or doctor --kill-orphans before the user says so.
  Tell the user; if the user closes that Excel, vbaloop recovery clear; if they say it may be discarded, vbaloop session stop --discard
- busy (exit 3): Excel is in cell edit mode or showing a dialog - retry the same command shortly
- If the user already has the canonical workbook open in Excel (read-only refusal hint): after the user confirms,
  vbaloop session attach - push/test then work in that Excel and session stop only detaches
- Tests are Public Sub Test_*() (no arguments); assert with LogShim.AssertEq (no Debug.Assert)
- src\LogShim.bas is a vendored copy - do not edit directly; update with vbaloop init
- When push fails, the hint states the workbook state (untouched / fully restored / discarded / session poisoned - not saved).
  After a failed push the canonical workbook on disk is always in its pre-push state.
  On a compile error (payload.compile.result = compile_error), fix it and push again.
  Exception: if the error is pre-existing and unrelated to this change, ask the user; only after they confirm, push --skip-compile.
  If you fixed something but the same error repeats and push keeps returning cached:true, push --force bypasses the fingerprint cache.
  To revert the workbook file itself, vbaloop rollback --latest (list: rollback --list;
  exit 3 while a persistent session is alive - session stop first)
- In ephemeral mode test/run open the canonical workbook read-only - sheets changed by tests are not saved
- Distribution artifact: vbaloop dist --out <path> (clean copy build - filter/stamp/sanitize/VBE compile check/optional bootstrap).
  push must come first (refused when last push != src). Procedure/failure playbook: SKILL.md
- pull is blocked when there are unsynced src edits - by default apply them with push first;
  only when the workbook is the truth, overwrite with pull --force
- No direct COM access to the session Excel (GetActiveObject, New-Object -ComObject etc. - including reads/diagnostics).
  To see what is really in the workbook, vbaloop diff --json (read-only; --module <name> dumps a single module)
- src conventions: UTF-8 without BOM + CRLF + exactly one trailing newline. Characters outside CP949 (em dash U+2014, emoji) are refused by push
- If you find a defect in vbaloop itself: create a standard report with vbaloop bugreport "<title>" and fill it in
  (procedure: .claude\skills\vbaloop\SKILL.md)
- Detailed procedures/failure playbooks: .claude\skills\vbaloop\SKILL.md (managed by vbaloop init - do not edit directly)
<!-- vbaloop:managed:end -->
