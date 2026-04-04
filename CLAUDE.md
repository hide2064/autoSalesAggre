# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

`autoSalesAggre.xlsm` — an Excel VBA macro workbook that:
- Reads multiple TSV files with mismatched headers via a Config name-normalization table
- Consolidates rows into an "all" sheet with master lookups (product name, commission)
- Provides department- and date-filtered hierarchical aggregation in a 集計 sheet

## Setup (first time)

1. In Excel: **File → Options → Trust Center → Trust Center Settings → Macro Settings** → check "Trust access to the VBA project object model"
2. Run: `cscript setup\create_workbook.vbs`
3. Open `autoSalesAggre.xlsm` and enable macros

## Rebuild after src/ changes

Delete `autoSalesAggre.xlsm`, then re-run `cscript setup\create_workbook.vbs`.

## Module responsibilities

| Module | Responsibility |
|--------|---------------|
| `modConfig` | All constants (column indices, sheet/cell addresses, HDR_* header names) + `NewDict()` + `LoadProductDict`, `LoadCommissionDict`, `LoadHeaderMap`, `RefreshDeptList` |
| `modFileIO` | `SelectFiles` (GetOpenFilename) + `LoadTsvToSheet` (two-pass bulk array write, all-text format) |
| `modDataProcess` | `BuildAllSheet` (header map + master lookup + bulk write) + `CollectUniqueDepts` |
| `modAggregation` | `Rebuild` (filter allData → dictSummary → `DrawAggrTable`) |
| `modUIControl` | `RunAll` (orchestrates all modules) + `LogMessage` (callable from any module) |
| `modSetup` | One-time init: sheet naming/layout + `InjectAggrEvent` (adds Worksheet_Change to 集計 sheet module) |

## Config sheet layout

| Column | Content |
|--------|---------|
| A–B | 製品マスタ: 製品コード → 製品名 (from row 3) |
| D–E | 口銭マスタ: 売上種別 → 口銭比率% (from row 3) |
| G–H | ヘッダー名寄せ: 正規名 → カンマ区切りエイリアス (from row 3) |
| J | 集計用部署リスト: J2="全部署" (fixed), J3+ auto-updated after each RunAll |

## Key design decisions

- `NewDict()` in modConfig creates all Scripting.Dictionary objects (case-insensitive, consistent)
- `HDR_*` constants in modConfig are the single source of truth for all column header strings — used in BuildAllSheet header writes, ProcessSourceSheet Select Case mapping, and modSetup sample data
- `LogMessage` is Public in modUIControl so it can be called unqualified from modDataProcess
- `Application.EnableEvents = False` is set during `RunAll` to prevent Worksheet_Change firing mid-process; re-enabled before calling `Rebuild` at the end
- `dictSummary` key format: `製品名 & "||" & 客先名` — the `||` separator avoids collisions with normal text
- TSV data is loaded in two passes: first to find dimensions, then bulk-written as a 2D Variant array with `NumberFormat = "@"` set on the range first to preserve leading zeros
- `ClearAggrTable` is called AFTER the no-data guard in `Rebuild` — if the all sheet is empty, the aggregation view is preserved rather than blanked
