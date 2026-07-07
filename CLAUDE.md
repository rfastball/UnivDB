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

## 금지 사항 (설계 결정으로 확정됨)

- Core에 DB 의미론 주입 금지 (`Core_JoinTables`, `Core_GroupBy` 류)
- 쿼리 실행 경로에 `someDict.Keys` / `New Collection` / `clsSet` 재도입 금지
- raw `Dictionary(key -> Collection)` 인덱스 경로 재도입 금지 (인덱스 진실은 `DbHashIndex` 하나)
- public 호환 shim(`SelectRows` 등) 시그니처 변경 금지
- 백엔드 자동 선택 도입 금지 - 기본은 DictHashIndex (policy v2, 2026-07-07 재확정),
  Core는 build/join/cursor/GroupBy-heavy 명시 옵션 + 벤치 기반 수동 판단

## 알려진 미결 사항

- GroupBy 키 의미론은 인덱스 계약과 통일됨 (R1/D2 결정, 2026-07-07). 단일 컬럼은 인덱스 커서
  기반 엔진화 (R2), 복합 키는 길이 접두 인코딩으로 충돌-프리 (R3). 복합 경로의 누적기 엔진화와
  hot column cache(R4 보류)는 병목 관측 시 재평가
- NULL/타입 의미론(비결정 D)은 1.0 전 확정 필요 - GroupBy 그룹 컬럼 값의 String화도 이 소관

## 설계 문서 (읽는 순서)

- `src/docs/tasks/2026-07-05_next_direction.md` - 현재 상태·실행 계획 (Phase R0~R5). 작업 착수 전 필독
- `src/docs/UnivDB_design_decisions_202607.md` - 세대 계약·스토리지 레이아웃 결정과 근거
- `src/docs/UnivDB_refactor_plan.md` - 내부 실행 표현 전환 원 설계 (Phase 0~8, 완료된 이력)
