# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

`autoSalesAggre.xlsm` — an Excel VBA macro workbook that:
- Reads TSV / CSV / Excel files with mismatched headers via a Config name-normalization table
- Consolidates rows into an "all" sheet with master lookups (product name, commission)
- Provides department- and date-filtered hierarchical aggregation in a 集計 sheet
- Uploads data to SharePoint via Power Automate HTTP trigger (separate URLs for 集計 / all)
- Generates a clustered column chart from aggregation results
- Provides an Excel native PivotTable on a dedicated ピボット sheet
- Generates a monthly summary in a 月次サマリー sheet
- Reports data errors (unknown codes, duplicate rows) in an エラー sheet
- Supports non-interactive batch execution via `setup\run_automate.vbs`

## Setup (first time)

1. In Excel: **File → Options → Trust Center → Trust Center Settings → Macro Settings** → check "Trust access to the VBA project object model"
2. Run: `cscript setup\create_workbook.vbs`
3. Open `autoSalesAggre.xlsm` and enable macros

## Rebuild after src/ changes

Delete `autoSalesAggre.xlsm`, then re-run `cscript setup\create_workbook.vbs`.

## Headless / batch execution

```
cscript setup\run_automate.vbs <フォルダパス>
```
Processes all TSV/CSV/xlsx files in the specified folder without showing any dialogs. Suitable for Windows Task Scheduler.

## Module responsibilities

| Module | Responsibility |
|--------|---------------|
| `modConfig` | All constants (column indices, sheet/cell addresses, HDR_*, CLR_*) + `NewDict()` + `LoadProductDict`, `LoadCommissionDict`, `LoadHeaderMap`, `RefreshDeptList`, `LoadPowerAutomateUrl`, `LoadPowerAutomateUrlAll`, `ValidateConfig` |
| `modFileIO` | `SelectFiles` + `LoadFileToSheet` (dispatcher) + `LoadDelimitedToSheet` (TSV/CSV) + `LoadXlsxSheetToSheet` (Excel) |
| `modDataProcess` | `BuildAllSheet` (header map + master lookup + dedup + bulk write) + `CollectUniqueDepts` |
| `modAggregation` | `Rebuild` (filter → dictSummary → `DrawAggrTable`) + `SaveFilter` / `RestoreFilter` |
| `modUIControl` | `RunAll` (orchestrates all modules) + `RunAllHeadless` (batch) + `LogMessage` |
| `modSharePoint` | `UploadToSharePoint` (集計 sheet) + `UploadAllToSharePoint` (all sheet, uses M3 URL) |
| `modChart` | `DrawAggrChart` (clustered column chart from 集計 sheet parent rows) |
| `modPivot` | `BuildPivot` (creates/refreshes PivotTable on ピボット sheet) |
| `modError` | `ClearErrorSheet` / `LogError` / `GetErrorCount` / `ActivateErrorSheet` |
| `modExport` | `ExportAggrToFile` (exports 集計 sheet to .xlsx) |
| `modMonthly` | `BuildMonthly` (month-by-month summary on 月次サマリー sheet) |
| `modSetup` | One-time init: sheet naming/layout/buttons + `InjectAggrEvent` |

## Sheets

| Sheet | Purpose |
|-------|---------|
| `main` | Execution log (timestamp + message) + "ファイルを読み込む" + "エラーを確認する" buttons |
| `Config` | Master data: 製品マスタ (A–B), 口銭マスタ (D–E), ヘッダー名寄せ (G–H), 部署リスト (J), PA URLs (M2/M3), フィルター保存 (O) |
| `all` | Normalized consolidated data (11 columns, row 2 onwards) |
| `集計` | Filtered hierarchical aggregation + chart / export / SharePoint / filter save/restore buttons |
| `ピボット` | Excel native PivotTable (auto-updated by RunAll, manually by button) |
| `エラー` | Data processing errors: unknown codes, duplicate rows (cleared each RunAll) |
| `月次サマリー` | Month-by-month sales totals (auto-updated by RunAll, manually by button) |

## Config sheet layout

| Column | Content |
|--------|---------|
| A–B | 製品マスタ: 製品コード → 製品名 (from row 3) |
| D–E | 口銭マスタ: 売上種別 → 口銭比率% (from row 3) |
| G–H | ヘッダー名寄せ: 正規名 → カンマ区切りエイリアス (from row 3) |
| J | 集計用部署リスト: J2="全部署" (fixed), J3+ auto-updated after each RunAll |
| L–M | SharePoint連携: M2=集計送信URL, M3=全データ送信URL (M3未設定時はM2にフォールバック) |
| O–P | フィルター条件保存: O2=部署, O3=開始日, O4=終了日 |

## Key constants (modConfig)

### Sheet names
- `SH_MAIN`, `SH_CONFIG`, `SH_ALL`, `SH_AGGR` ("集計"), `SH_PIVOT` ("ピボット")
- `SH_ERROR` ("エラー"), `SH_MONTHLY` ("月次サマリー")

### Column indices (all sheet)
- `ALL_COL_CLIENT`(1) through `ALL_COL_SOURCE`(11), `ALL_TOTAL_COLS`=11

### Header strings
- `HDR_CLIENT`, `HDR_PROD_CODE`, `HDR_AMOUNT`, `HDR_UNIT_PRICE`, `HDR_QTY`, `HDR_DATE`, `HDR_SALE_TYPE`, `HDR_DEPT`, `HDR_PROD_NAME`, `HDR_MARGIN`, `HDR_SOURCE`

### Color constants (Long values, pre-computed as R + G×256 + B×65536)
- `CLR_HEADER_BG` = RGB(200,220,240) — header row background (blue-tinted)
- `CLR_GROUP_ROW` = RGB(220,220,220) — aggregation group row background (grey)
- `CLR_CHART_AMT` = RGB(70,130,180) — chart series 1: 売上金額合計 (steel blue)
- `CLR_CHART_MARGIN` = RGB(255,165,0) — chart series 2: 口銭総額 (orange)
- `CLR_PLOT_AREA` = RGB(248,248,248) — chart plot area background (near-white)
- `CLR_LABEL_TEXT` = RGB(100,100,100) — descriptive label text (grey)
- `CLR_ERROR_ROW` = RGB(255,220,220) — error sheet row background (light red)
- `CLR_MONTHLY_HDR` = RGB(200,240,220) — monthly sheet header background (light green)

### Other constants
- `DICT_KEY_SEP = "||"` — separator in dictSummary keys and dedup keys
- `COL_MAP_COUNT = 8` — size of colMap array in ProcessSourceSheet
- `CFG_PA_URL_ROW = 2`, `CFG_PA_URL_ALL_ROW = 3` — PA URL rows in Config M column
- `CFG_SAVED_FILTER_COL = 15` (O列), rows 2/3/4 — saved filter state storage

## Key design decisions

- `NewDict()` in modConfig creates all Scripting.Dictionary objects (case-insensitive, consistent)
- `HDR_*` constants are the single source of truth for all column header strings
- `CLR_*` constants are the single source of truth for all colors; no `RGB()` calls outside modConfig
- `LogMessage` is Public in modUIControl so it can be called unqualified from any module
- `LogError` in modError writes to the エラー sheet; called from modDataProcess for unknown codes and duplicate rows
- Duplicate row detection: key = `sourceFile & DICT_KEY_SEP & all 8 TSV fields`; rows with the same key within a run are skipped (protects against accidental double-selection of the same file)
- `Application.EnableEvents = False` is set during `RunAll` to prevent Worksheet_Change firing mid-process
- `dictSummary` key format: `製品名 & DICT_KEY_SEP & 客先名`
- TSV/CSV data is loaded in two passes: first to find dimensions, then bulk-written as a 2D Variant array
- Excel (.xlsx) input: source workbook opened read-only, data extracted via Variant array, workbook closed before writing to destination sheet
- Filter persistence: `SaveFilter` stores B1/B2/B3 to Config O2/O3/O4; called automatically after each Rebuild; `RestoreFilter` is the button-facing counterpart
- PA URLs: M2 = 集計シート送信用, M3 = allシート全データ送信用; `LoadPowerAutomateUrlAll()` falls back to M2 if M3 is empty
- `ValidateConfig()` checks product/commission/header map counts and commission rate ranges; logs warnings via `LogMessage`; returns issue count (0 = OK)
- `RunAllHeadless(folderPath)` enumerates files via `Dir()`, processes without any dialogs, saves workbook on completion
- `create_workbook.vbs` uses `On Error Resume Next` throughout with explicit `xlApp.Quit` on every failure path
- PivotTable (modPivot): `ChangePivotCache` + refresh on update; create + `ConfigurePivotTable` on first run
- `SendHttpPost` in modSharePoint is a `Private Function` returning HTTP status as `Long`
