# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

UnivDB is a VBA in-memory database library for Excel. The host workbook lives at `../dev_UnivDB.xlsb` (parent directory); the `.bas`/`.cls` files in this directory are the exported source that gets re-imported into the VBA project.

The library's full public API is documented in `UnivDB_Reference.md` (v2.5.1). Read it before writing UnivDB-consuming code — it is the canonical reference for `UnivDB.*` facade calls, `clsQuery`/`clsReportBuilder`/`clsTableBuilder` chains, JoinType enums, and end-to-end pipeline patterns.

## Project-Wide Conventions (inherited)

`../../CLAUDE.md` defines the workspace-wide module/procedure header schema. All new modules must include `Attribute VB_Name`, then `' @desc ...` / `' @depends ...` lines, then `Option Explicit`. Public Subs/Functions get a one-line `' @desc` comment immediately above the declaration (40/60 char limits, no "관리"/"처리"/"담당" alone). Private procedures don't get `@desc`.

## Architecture

**Pattern:** Facade (`UnivDB` module) + Method Chaining + Copy-on-Write for Report pipelines.

**Layered structure** — understanding these layers is required before touching join/query code:

1. **`clsTable`** — data container. 2D Variant array + lazy column indexes (`m_dicIndices`). Public surface includes `SelectRows`/`SelectRowsByIndex`, `GetUniqueKeys`, `GroupBy`, `Compute`, `UpsertFrom`, `OrderBy`. Indexes are built on demand and cached per column.
2. **`clsSet`** — high-speed set engine (Dictionary-backed, case-insensitive). Powers join classification via `Compare(..., srAllSplit)` returning `(intersect, leftOnly, rightOnly)` in a single pass.
3. **`modEngine.JoinTable`** — universal join (Inner/Left/Right/Full/Anti/Semi). Two-stage: (a) classify keys via `clsSet`, (b) size result, (c) materialize rows. v2.5.1 hoists key-column indexes out of inner loops — preserve that pattern when editing.
4. **`clsQuery`** — chainable CRUD. Order of evaluation matters: `.Where()` (index, O(1)) should precede `.Filter()` (scan); both narrow `m_setCandidates` via intersection.
5. **`clsReportBuilder`** — ETL pipeline with **Copy-on-Write**. `Init()` references the source table by ref; the first structural/value mutation calls `EnsureSafeCopy()` to clone via `LoadFromArray ..., True`. Never mutate `m_tWork` directly without going through `EnsureSafeCopy` — that would corrupt the registered source table.
6. **`clsTableBuilder`** — row-by-row assembly with a cursor (`m_sCursor`/`m_vCursor`). Chaining requires `Call` (see Reference §6).
7. **`UnivDB.bas`** — facade. `ResolveTable(vInput)` accepts either a registered table name (`String`) or a `clsTable` object; all facade entry points use it, so consumer code can mix names and objects freely.

**Data flow:** Sheet/Array → `UnivDB.RegisterTable` → in-memory `clsTable` (stored in `m_dicTables` keyed by `UCase(Trim(name))`) → `Query`/`Report`/`Join` → `clsTable` result → `.DumpToSheet` / `.BuildToSheet`.

## Critical Lifecycle Rule

Any sub that calls `UnivDB.InitDB` **must** call `UnivDB.ClearDB` in a `Finally` block. `clsTable` instances must not outlive the entry-point sub. The standard pattern is `On Error GoTo ErrHandler` → work → `Finally: UnivDB.ClearDB : Exit Sub` → `ErrHandler: modLogger.LogError ... : Resume Finally`. See Reference §8.4.

## Conventions That Affect Edits

- **Header lookups are case-insensitive.** `clsTable` stores headers as `UCase(Trim(name))`; all engine code calls `tLeft.Headers(UCase(sLeftKey))`. Don't compare raw user-supplied column names against `m_dicHeaders.Keys` without uppercasing.
- **Join Prefix safety.** Joined columns get a prefix (e.g. `"Proj_"`) and are intentionally invisible to `UpsertFrom` because `tSource.headers.exists` filters them out. The Reference flags this as "do not modify Upsert" — respect it.
- **Variable prefixes** (used throughout the codebase, follow them): `t` clsTable, `o` Object, `s` String, `b` Boolean, `v` Variant, `l` Long, `dic` Dictionary, `col` Collection, `rng` Range.
- **CRUD return values.** `.Insert/.Update/.Delete` return `Long` (affected row count); they terminate a chain — don't `Set` their result.
- **Logging.** `modLogger.Log` is gated by a module-level `ENABLE_ALL_LOGGING` constant (currently `False`). `LogError` is always-on. Don't sprinkle `Debug.Print` — route through `modLogger`.

## Testing

Tests are VBA modules named `Test*.bas` (`TestJoin`, `TestQuery`, `TestFilter`, `TestTableBuilder`, `TestRefactor_v2_5`). Each exposes a `RunAll_*_Tests` Public Sub that prints pass/fail to the Immediate window and a sentinel banner on completion. To run: open `../dev_UnivDB.xlsb` in Excel, Alt+F11, F5 the runner — there is no headless test harness. Tests create and tear down their own scratch sheet (e.g. `TEST_SHEET = "TestJoinSheet"`).

## Build & Sync

There is no build step. The `.bas`/`.cls` files in this directory are the source-of-truth exports. The `archive/` directory holds timestamped snapshots of prior exports (format `YYYYMMDD_HHMMSS`); treat them as read-only history.

When editing VBA source, edit the `.bas`/`.cls` here — the user re-imports into `dev_UnivDB.xlsb` manually. Preserve the class header block (`VERSION 1.0 CLASS` … `Attribute VB_Exposed = True`) verbatim on `.cls` files; corrupting it breaks the import.

## Files To Read First For Common Tasks

- **Adding a query operator:** `clsQuery.cls` (`Filter`/`Where` show the intersection pattern), `clsSet.cls`.
- **Changing join semantics:** `modEngine.bas` (`JoinTable`), `clsSet.cls` (`Compare`/`srAllSplit`).
- **Adding a report transform:** `clsReportBuilder.cls` — always call `EnsureSafeCopy` first.
- **Adding a `clsTable` operation:** `clsTable.cls` — note the lazy-index cache (`m_dicIndices`); invalidate it if you mutate column structure.
- **End-to-end usage examples:** `modTemplate.bas` and Reference §8.4.
