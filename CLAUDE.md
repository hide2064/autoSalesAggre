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

## Encoding rules

All `.bas` files (`src/`) and `.vbs` files (`setup/`) must be saved in **Shift-JIS (CP932)**. After editing, convert with:

```python
python3 -c "
import os, glob
for fpath in glob.glob('src/*.bas') + glob.glob('setup/*.vbs'):
    with open(fpath, 'r', encoding='utf-8') as f:
        text = f.read()
    with open(fpath, 'w', encoding='cp932') as f:
        f.write(text)
"
```

## Module responsibilities

| Module | Responsibility |
|--------|---------------|
| `modConfig` | 全定数 (HDR_*, CFG_*, SH_*, AGGR_*) + `NewDict()` + マスタ読込関数群 (`LoadProductDict`, `LoadCommissionDict`, `LoadHeaderMap`, `LoadAllColDef`, `GetAllColIndex`, `RefreshDeptList`) |
| `modFileIO` | `SelectFiles` (GetOpenFilename) + `LoadTsvToSheet` (two-pass bulk array write, all-text format) |
| `modDataProcess` | `BuildAllSheet` (動的列定義 + マスタルックアップ + バルク書き込み) + `CollectUniqueDepts` |
| `modAggregation` | `Rebuild` (フィルタ → dictSummary → `DrawAggrTable`) |
| `modUIControl` | `RunAll` (全モジュールのオーケストレーション) + `LogMessage` (全モジュールから呼び出し可) |
| `modSetup` | 初回セットアップ: シート作成・レイアウト + `InjectAggrEvent` (集計シートモジュールに Worksheet_Change を注入) |
| `modSharePoint` | `UploadToSharePoint` (集計シート) + `UploadAllToSharePoint` (allシート) + `SendJson` (共通HTTP送信) |
| `modChart` | `DrawAggrChart` (集計シートのグラフ作成) |

## Config sheet layout

| Column | Content |
|--------|---------|
| A–B | 製品マスタ: 製品コード → 製品名 (from row 3) |
| D–E | 口銭マスタ: 売上種別 → 口銭比率% (from row 3) |
| G–I | ヘッダー名寄せ: G=正規名 / H=カンマ区切りエイリアス / I=Allシート列名 (from row 3) |
| J | 集計用部署リスト: J2="全部署" (fixed), J3+ auto-updated after each RunAll |

**Allシート列の制御:**
- Config G〜I列の名寄せテーブルで Allシートの列構成を定義する
- I列（Allシート列名）に値がある行だけ、その行の並び順で Allシートに出力される
- I列が空白の行は Allシートに出力されない
- 計算列（製品名・口銭按分）とソースファイル名は常に末尾3列に固定出力される

## LogMessage rules

- Do **not** start log strings with `=` or `-` — Excel interprets them as formulas and throws an error
- Use `【】` brackets for section markers (e.g. `"【処理開始】"`)

## Key design decisions

- `NewDict()` in modConfig creates all Scripting.Dictionary objects (case-insensitive, consistent)
- `HDR_*` constants in modConfig are the single source of truth for all column header strings
- `AGGR_INDENT` (全角スペース2文字) is shared between `modAggregation` and `modChart` — both must use this constant for child-row detection; changing the indent string requires updating only this one place
- `AGGR_KEY_SEP = "||"` is the separator for dictSummary keys (`製品名 & AGGR_KEY_SEP & 客先名`) — defined in modConfig
- `LogMessage` is Public in modUIControl so it can be called unqualified from any module
- `Application.EnableEvents = False` is set during `RunAll` to prevent Worksheet_Change firing mid-process; re-enabled before calling `Rebuild` at the end
- `ClearAggrTable` is called AFTER the no-data guard in `Rebuild` — if the all sheet is empty, the aggregation view is preserved rather than blanked
- TSV data is loaded in two passes: first to find dimensions, then bulk-written as a 2D Variant array with `NumberFormat = "@"` set on the range first to preserve leading zeros
- `LoadAllColDef()` reads Config G〜I columns in one pass and returns an ordered Dictionary (canonical name → All sheet column name); insertion order = column order in All sheet
- `GetAllColIndex(wsAll, headerName)` scans All sheet row 1 to find a column by name — called at the top of `Rebuild` and `UploadAllToSharePoint` to resolve column positions dynamically
- `InjectAggrEvent` injects VBA code as a string — any modConfig constant names referenced inside that string must be kept in sync manually (see warning comment in modSetup)
