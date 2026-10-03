# vbaflow: xlam session adopted but run cannot locate workbook

- Date: 2026-10-03
- Project: UnivDB
- Status: open
- Severity: P2 (artifact verification tooling)
- Artifact: dist/UnivDB.xlam, built by vbaflow dist 2.8.1

## Reproduction

Use a separate vbaflow config with excel.workbook pointing to the absolute path
of the generated .xlam, then run:

    vbaflow run UnivDB.LibVersion --config _dbg/release-verify/vbaflow.toml --json
    vbaflow session start --config _dbg/release-verify/vbaflow.toml --json
    vbaflow run UnivDB.LibVersion --config _dbg/release-verify/vbaflow.toml --json

The first run starts Excel successfully. session start reports action=adopted.
Both run calls return exitCode=2, hint="session start", and:

    세션에 대상 워크북이 열려 있지 않음: ...\dist\UnivDB.xlam

No direct COM access was used. The tool-created verification session was stopped
with --no-save. vbaflow dist --bootstrap UnivDB.InitDB succeeded before this check;
source tests passed 122/122. This does not prove a defect in the generated add-in.

## Expected

run should find the add-in it opened, or explicitly reject unsupported .xlam
sessions without the misleading session-start retry hint.

## Follow-up

Inspect workbook discovery for hidden add-ins. Cause is not yet confirmed.
The installed vbaflow bugreport --list command itself returned:
"_VbaFlow 레포를 찾지 못함 (exe 위치에서 상향 탐색 실패)"
and requested a project-local memo, so this report is stored here.
