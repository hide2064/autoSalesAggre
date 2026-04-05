# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

`autoSalesAggre.xlsm` — an Excel VBA macro workbook that:
- Reads multiple TSV files with mismatched headers via a Config name-normalization table
- Consolidates rows into an "all" sheet with master lookups (product name, commission)
- Provides department- and date-filtered hierarchical aggregation in a 集計 sheet
- Uploads data to SharePoint via Power Automate HTTP trigger
- Generates a clustered column chart from aggregation results
- Provides an Excel native PivotTable on a dedicated ピボット sheet

## Setup (first time)

1. In Excel: **File → Options → Trust Center → Trust Center Settings → Macro Settings** → check "Trust access to the VBA project object model"
2. Run: `cscript setup\create_workbook.vbs`
3. Open `autoSalesAggre.xlsm` and enable macros

## Rebuild after src/ changes

Delete `autoSalesAggre.xlsm`, then re-run `cscript setup\create_workbook.vbs`.

## Module responsibilities

| Module | Responsibility |
|--------|---------------|
| `modConfig` | All constants (column indices, sheet/cell addresses, HDR_* header names, CLR_* color constants) + `NewDict()` + `LoadProductDict`, `LoadCommissionDict`, `LoadHeaderMap`, `RefreshDeptList`, `LoadPowerAutomateUrl` |
| `modFileIO` | `SelectFiles` (GetOpenFilename) + `LoadTsvToSheet` (two-pass bulk array write, all-text format) |
| `modDataProcess` | `BuildAllSheet` (header map + master lookup + bulk write) + `CollectUniqueDepts` |
| `modAggregation` | `Rebuild` (filter allData → dictSummary → `DrawAggrTable`) |
| `modUIControl` | `RunAll` (orchestrates all modules) + `LogMessage` (callable from any module) |
| `modSharePoint` | `UploadToSharePoint` (集計 sheet → Power Automate) + `UploadAllToSharePoint` (all sheet → Power Automate) |
| `modChart` | `DrawAggrChart` (clustered column chart from 集計 sheet parent rows) |
| `modPivot` | `BuildPivot` (creates/refreshes Excel native PivotTable on ピボット sheet from all sheet data) |
| `modSetup` | One-time init: sheet naming/layout + `InjectAggrEvent` (adds Worksheet_Change to 集計 sheet module) |

## Sheets

| Sheet | Purpose |
|-------|---------|
| `main` | Execution log (timestamp + message) + "ファイルを読み込む" button |
| `Config` | Master data: 製品マスタ (A–B), 口銭マスタ (D–E), ヘッダー名寄せ (G–H), 部署リスト (J), Power Automate URL (M2) |
| `all` | Normalized consolidated data (11 columns, 2 row onwards) |
| `集計` | Filtered hierarchical aggregation + chart button + SharePoint button |
| `ピボット` | Excel native PivotTable (auto-updated by RunAll, manually by button) |

## Config sheet layout

| Column | Content |
|--------|---------|
| A–B | 製品マスタ: 製品コード → 製品名 (from row 3) |
| D–E | 口銭マスタ: 売上種別 → 口銭比率% (from row 3) |
| G–H | ヘッダー名寄せ: 正規名 → カンマ区切りエイリアス (from row 3) |
| J | 集計用部署リスト: J2="全部署" (fixed), J3+ auto-updated after each RunAll |
| L–M | SharePoint連携: L2=ラベル, M2=Power Automate HTTP trigger URL |

## Key constants (modConfig)

### Sheet names
- `SH_MAIN`, `SH_CONFIG`, `SH_ALL`, `SH_AGGR` ("集計"), `SH_PIVOT` ("ピボット")

### Column indices (all sheet)
- `ALL_COL_CLIENT`(1) through `ALL_COL_SOURCE`(11), `ALL_TOTAL_COLS`=11

### Header strings
- `HDR_CLIENT`, `HDR_PROD_CODE`, `HDR_AMOUNT`, `HDR_UNIT_PRICE`, `HDR_QTY`, `HDR_DATE`, `HDR_SALE_TYPE`, `HDR_DEPT`, `HDR_PROD_NAME`, `HDR_MARGIN`, `HDR_SOURCE`

### Color constants (Long values, pre-computed)
- `CLR_HEADER_BG` = RGB(200,220,240) — header row background (blue-tinted)
- `CLR_GROUP_ROW` = RGB(220,220,220) — aggregation group row background (grey)
- `CLR_CHART_AMT` = RGB(70,130,180) — chart series 1: 売上金額合計 (steel blue)
- `CLR_CHART_MARGIN` = RGB(255,165,0) — chart series 2: 口銭総額 (orange)
- `CLR_PLOT_AREA` = RGB(248,248,248) — chart plot area background (near-white)
- `CLR_LABEL_TEXT` = RGB(100,100,100) — descriptive label text (grey)

### Other constants
- `DICT_KEY_SEP = "||"` — separator in dictSummary keys (`製品名 & DICT_KEY_SEP & 客先名`)
- `COL_MAP_COUNT = 8` — size of colMap array in ProcessSourceSheet (ALL_COL_CLIENT to ALL_COL_DEPT)
- `CFG_PA_URL_ROW = 2`, `CFG_PA_URL_COL = 13` — Power Automate URL cell (M2)

## Key design decisions

- `NewDict()` in modConfig creates all Scripting.Dictionary objects (case-insensitive, consistent)
- `HDR_*` constants in modConfig are the single source of truth for all column header strings — used in BuildAllSheet header writes, ProcessSourceSheet Select Case mapping, and modSetup sample data
- `CLR_*` constants in modConfig are the single source of truth for all colors — used in modSetup (sheet headers), modAggregation (group rows), modChart (series colors); no `RGB()` calls outside modConfig
- `LogMessage` is Public in modUIControl so it can be called unqualified from modDataProcess
- `Application.EnableEvents = False` is set during `RunAll` to prevent Worksheet_Change firing mid-process; re-enabled before calling `Rebuild` at the end
- `dictSummary` key format: `製品名 & DICT_KEY_SEP & 客先名` — the `||` separator avoids collisions with normal text; defined as `DICT_KEY_SEP` constant to avoid scattered string literals
- TSV data is loaded in two passes: first to find dimensions, then bulk-written as a 2D Variant array with `NumberFormat = "@"` set on the range first to preserve leading zeros
- `ClearAggrTable` is called AFTER the no-data guard in `Rebuild` — if the all sheet is empty, the aggregation view is preserved rather than blanked
- `colMap(COL_MAP_COUNT - 1)` in ProcessSourceSheet maps all-sheet column indices (0-based) to source column numbers; sized by `COL_MAP_COUNT` constant rather than a hardcoded literal
- PivotTable (modPivot): checks for existing table by iterating `wsPivot.PivotTables`; if found → `ChangePivotCache` + refresh; if not → create + configure. This avoids name collisions on repeated RunAll calls.
- `SendHttpPost` in modSharePoint is a `Private Function` returning HTTP status as `Long` (200/202=success, -1=exception). Both upload subs use `Select Case` on the return value for clean error handling.
- `create_workbook.vbs` uses `On Error Resume Next` throughout the Excel operation block with explicit `xlApp.Quit` on every failure path to prevent orphaned Excel processes.
