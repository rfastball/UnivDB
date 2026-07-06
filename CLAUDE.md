# AGENTS.md / CLAUDE.md - UnivDB

(하네스 루프 계약은 상위 Project/AGENTS.md 3장 참조)

최초 1회: ..\..\_VbaHarness\vh install (사용자 PATH 등록 - 이후 어디서든 bare vh)
일상 루프(src\ 기준, 설치 후):
- 실행: vh run  (설치 전엔 ..\..\_VbaHarness\vh run)
- 실패 로그: vh log -FailedOnly
- 커밋 전 검사: vh check

하네스 대상 워크북: dev_UnivDB.xlsb

공용 패키지 의존성: UnivCore (deps.lock으로 버전 고정)
- src\Core_* 는 설치 산출물 (gitignore, 직접 수정 금지 - 수정은 UnivCore에서)
- 드리프트/구버전 검사: vh deps check, 재설치: vh deps install, 최신 재고정: vh deps update
- vh run이 lock 기준 자동 hydrate하므로 평소엔 신경 쓸 필요 없음
