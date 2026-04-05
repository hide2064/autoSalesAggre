Attribute VB_Name = "modPivot"
Option Explicit

' ============================================================
' modPivot — ピボットテーブル作成・更新モジュール
'
' 役割:
'   ・all シートのデータを元に「ピボット」シートへ Excel ネイティブの
'     PivotTable を作成・更新する BuildPivot を提供する。
'   ・RunAll 完了後に自動呼び出しされるほか、ピボットシートの
'     「ピボットテーブル更新」ボタンからも手動実行できる。
'
' デフォルトのフィールド構成:
'   行 (Rows)      : 製品名 → 客先名（2階層）
'   フィルター     : 部署、売上種別（ドロップダウンで絞り込み可）
'   値 (Values)    : 売上金額合計（Sum）、売上数量合計（Sum）、部署取り分合計（Sum）
'   列 (Columns)   : なし（ユーザーが自由にフィールドをドラッグして追加可）
'
' 更新ロジック:
'   ・"SalesPivot" という名前の PivotTable が既に存在する場合は
'     PivotCache のソース範囲を新しいデータ行数に合わせて更新し
'     RefreshTable で再集計する（フィールド配置はユーザー設定を保持）。
'   ・存在しない場合は新規作成して初期フィールド構成を適用する。
'
' 制約:
'   ・PivotTable 名 "SalesPivot" はこのモジュール内で一意に管理する。
'     重複作成は行わない。
' ============================================================

' ピボットテーブルの識別名（新規作成・更新検索に使用）
Private Const PIVOT_NAME As String = "SalesPivot"

' ピボットテーブルの配置開始行（1〜3行目を UI 領域として確保）
Private Const PIVOT_START_ROW As Integer = 4

' ============================================================
' BuildPivot — ピボットテーブルを作成または更新する（Public）
'
' 処理概要:
'   1. all シートのデータ有無を確認
'   2. ピボットシートを取得（なければ作成）
'   3. "SalesPivot" が既存ならソース範囲を更新して RefreshTable
'   4. 存在しなければ新規作成して ConfigurePivotTable で初期設定
' ============================================================
Public Sub BuildPivot()
    Dim wsAll As Worksheet
    Dim wsPivot As Worksheet
    Dim lastRow As Long
    Dim srcRange As String
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim ptFound As PivotTable
    Dim ptExists As Boolean

    Set wsAll = ThisWorkbook.Sheets(SH_ALL)
    lastRow = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        LogMessage "ピボットテーブル: allシートにデータがありません（スキップ）"
        Exit Sub
    End If

    Set wsPivot = GetPivotSheet()

    ' --- ソース範囲アドレスを組み立て（ヘッダー行 + データ行）---
    ' 絶対アドレスでシート名を含めることで PivotCache がシートを正しく参照する
    srcRange = "'" & SH_ALL & "'!" & _
               wsAll.Range(wsAll.Cells(1, 1), _
                           wsAll.Cells(lastRow, ALL_TOTAL_COLS)).Address(ReferenceStyle:=xlA1)

    ' --- 既存の "SalesPivot" を検索 ---
    ptExists = False
    For Each pt In wsPivot.PivotTables
        If pt.Name = PIVOT_NAME Then
            ptExists = True
            Set ptFound = pt
            Exit For
        End If
    Next pt

    If ptExists Then
        ' --- 既存ピボット: ソース範囲を更新してリフレッシュ ---
        ' ChangePivotCache でデータ行数の変化（追加・削除）に追従する。
        ' フィールド配置はユーザーの設定を保持するため再設定しない。
        ptFound.ChangePivotCache ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=srcRange)
        ' 古いデータのキャッシュ残留を防ぐ
        ptFound.PivotCache.MissingItemsLimit = xlMissingItemsNone
        ptFound.RefreshTable
        LogMessage "ピボットテーブルを更新しました (" & (lastRow - 1) & "行)"
    Else
        ' --- 新規ピボット: PivotCache を作成してテーブルを構築 ---
        Set pc = ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=srcRange)
        pc.MissingItemsLimit = xlMissingItemsNone

        Dim newPt As PivotTable
        Set newPt = pc.CreatePivotTable( _
            TableDestination:=wsPivot.Cells(PIVOT_START_ROW, 1), _
            TableName:=PIVOT_NAME)

        ConfigurePivotTable newPt
        LogMessage "ピボットテーブルを新規作成しました (" & (lastRow - 1) & "行)"
    End If
End Sub

' ============================================================
' ConfigurePivotTable — 新規 PivotTable の初期フィールド構成を設定する（プライベート）
'
' 引数:
'   pt — 設定対象の PivotTable オブジェクト
'
' 設定内容:
'   行      : 製品名(1階層目) → 客先名(2階層目)
'   フィルター: 部署（ドロップダウン位置1）, 売上種別（ドロップダウン位置2）
'   値      : 売上金額合計(Sum) / 売上数量合計(Sum) / 部署取り分合計(Sum)
'   書式    : 値フィールドを "#,##0" の千区切り表示に統一
'   スタイル : "PivotStyleMedium9"（青系の中程度スタイル）
'
' 列フィールドは意図的に設定しない。
' ユーザーがフィールドリストから 部署 などを列にドラッグして
' クロス集計を自由に構築できるようにするため。
' ============================================================
Private Sub ConfigurePivotTable(pt As PivotTable)
    With pt
        ' --- 行フィールド: 製品名 → 客先名 の2階層 ---
        With .PivotFields(HDR_PROD_NAME)
            .Orientation = xlRowField
            .Position    = 1
        End With
        With .PivotFields(HDR_CLIENT)
            .Orientation = xlRowField
            .Position    = 2
        End With

        ' --- フィルターフィールド: 部署・売上種別 ---
        ' ページフィールド（シート上部のドロップダウン）として配置
        With .PivotFields(HDR_DEPT)
            .Orientation = xlPageField
            .Position    = 1
        End With
        With .PivotFields(HDR_SALE_TYPE)
            .Orientation = xlPageField
            .Position    = 2
        End With

        ' --- 値フィールド: 合計3種 ---
        .AddDataField .PivotFields(HDR_AMOUNT), "売上金額合計", xlSum
        .AddDataField .PivotFields(HDR_QTY),    "売上数量合計", xlSum
        .AddDataField .PivotFields(HDR_MARGIN),  "部署取り分合計", xlSum

        ' --- 値フィールドの数値書式を千区切りに統一 ---
        .DataFields("売上金額合計").NumberFormat   = "#,##0"
        .DataFields("売上数量合計").NumberFormat   = "#,##0"
        .DataFields("部署取り分合計").NumberFormat = "#,##0"

        ' --- 表示設定 ---
        .TableStyle2  = "PivotStyleMedium9"  ' 青系中程度スタイル
        .RowGrand     = True                 ' 行の総計を表示
        .ColumnGrand  = True                 ' 列の総計を表示
        .ShowDrillIndicators = True          ' 展開/折りたたみインジケーターを表示

        ' コンパクト形式で表示（行フィールドを同一列に折りたたみ）
        .RowAxisLayout xlCompactRow
    End With
End Sub

' ============================================================
' GetPivotSheet — ピボットシートを取得する（プライベート）
'
' SH_PIVOT("ピボット")シートが既に存在すれば返す。
' 存在しない場合は 集計 シートの後ろに新規作成して返す。
' ※ 通常は create_workbook.vbs と InitWorkbook で事前に作成されるため
'   フォールバックとして用意している。
' ============================================================
Private Function GetPivotSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_PIVOT)
    On Error GoTo 0

    If ws Is Nothing Then
        ' 集計シートの後ろに追加
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(SH_AGGR))
        ws.Name = SH_PIVOT
    End If

    Set GetPivotSheet = ws
End Function
