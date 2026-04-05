Attribute VB_Name = "modSharePoint"
Option Explicit

' ============================================================
' modSharePoint — SharePoint アップロードモジュール
'
' 役割:
'   ・集計シートの集計結果を Power Automate 経由で SharePoint に
'     アップロードする UploadToSharePoint を提供する。
'   ・all シートの全データを Power Automate 経由で SharePoint に
'     アップロードする UploadAllToSharePoint を提供する。
'
' 仕組み:
'   1. Config シートの M2 から Power Automate HTTP 要求トリガーの URL を取得
'   2. 対象シートのデータを JSON 文字列に変換
'   3. プライベート関数 SendHttpPost で HTTP POST 送信
'   4. 戻り値の HTTP ステータスコードを元に成功/失敗を判定
'      200 / 202 → 正常終了（Power Automate は通常 202 Accepted を返す）
'      -1        → 通信例外
'      その他    → HTTP エラー
'
' 拡張方法:
'   ・新しいシートのアップロードを追加する場合は、
'     ① JSON 組み立てロジックを持つ Public Sub を追加
'     ② SendHttpPost を呼び出して戻り値で成否を判定
'     ③ modSetup でボタンの OnAction に新 Sub を設定
'
' 送信 JSON フォーマット（集計シート版）:
'   {
'     "dept":       "全部署",
'     "fromDate":   "2026/01/01",
'     "toDate":     "2026/03/31",
'     "uploadedAt": "2026/04/04 12:00:00",
'     "rows": [
'       {"name": "製品A",       "amount": 150000, "qty": 30, "margin": 15000},
'       {"name": "　　得意先X", "amount": 100000, "qty": 20, "margin": 10000},
'       {"name": "総合計",      "amount": 200000, "qty": 40, "margin": 20000}
'     ]
'   }
'
' 送信 JSON フォーマット（all シート版）:
'   {
'     "uploadedAt": "2026/04/04 12:00:00",
'     "rows": [
'       {
'         "client":   "客先A",   "prodCode": "P001",
'         "amount":   10000,     "unitPrice": 1000,
'         "qty":      10,        "date":      "2026/01/15",
'         "saleType": "直販",    "dept":      "営業部",
'         "prodName": "製品A",   "margin":    1000,
'         "source":   "jan.tsv"
'       }, ...
'     ]
'   }
' ============================================================

' ============================================================
' UploadToSharePoint — 集計シートのデータを Power Automate 経由でアップロード
'
' 集計テーブルの全行（製品グループ行・客先明細行・総合計行）を
' JSON 配列に変換して送信する。フィルタ条件（部署・期間）も
' JSON の上位フィールドとして含める。
' ============================================================
Public Sub UploadToSharePoint()
    Dim paUrl As String
    Dim wsAggr As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim rowsJson As String
    Dim sep As String
    Dim jsonBody As String
    Dim httpStatus As Long

    paUrl = LoadPowerAutomateUrl()
    If paUrl = "" Then
        MsgBox "ConfigシートのM2にPowerAutomate URLが設定されていません。", _
               vbExclamation, "設定エラー"
        Exit Sub
    End If

    Set wsAggr = ThisWorkbook.Sheets(SH_AGGR)
    lastRow = wsAggr.Cells(wsAggr.Rows.Count, 1).End(xlUp).Row

    If lastRow < AGGR_DATA_ROW Then
        MsgBox "集計データがありません。先にデータを集計してください。", _
               vbExclamation, "データなし"
        Exit Sub
    End If

    ' --- 集計テーブルの各行を JSON オブジェクトに変換して連結 ---
    rowsJson = ""
    sep = ""
    For r = AGGR_DATA_ROW To lastRow
        If Trim(CStr(wsAggr.Cells(r, 1).Value)) = "" Then GoTo NextRow

        rowsJson = rowsJson & sep & "{"
        rowsJson = rowsJson & """name"":"   & JsonString(wsAggr.Cells(r, 1).Value) & ","
        rowsJson = rowsJson & """amount"":"  & JsonNumber(wsAggr.Cells(r, 2).Value) & ","
        rowsJson = rowsJson & """qty"":"     & JsonNumber(wsAggr.Cells(r, 3).Value) & ","
        rowsJson = rowsJson & """margin"":"  & JsonNumber(wsAggr.Cells(r, 4).Value)
        rowsJson = rowsJson & "}"
        sep = ","
NextRow:
    Next r

    ' --- JSON ペイロード組み立て ---
    jsonBody = "{"
    jsonBody = jsonBody & """dept"":"       & JsonString(wsAggr.Range(AGGR_DEPT_CELL).Value) & ","
    jsonBody = jsonBody & """fromDate"":"   & JsonString(wsAggr.Range(AGGR_FROM_CELL).Value) & ","
    jsonBody = jsonBody & """toDate"":"     & JsonString(wsAggr.Range(AGGR_TO_CELL).Value)   & ","
    jsonBody = jsonBody & """uploadedAt"":" & JsonString(Format(Now(), "yyyy/mm/dd hh:mm:ss")) & ","
    jsonBody = jsonBody & """rows"":["      & rowsJson & "]"
    jsonBody = jsonBody & "}"

    ' --- HTTP POST 送信と結果処理 ---
    httpStatus = SendHttpPost(paUrl, jsonBody)
    Select Case httpStatus
        Case 200, 202
            LogMessage "集計シートSharePointアップロード完了 (HTTP " & httpStatus & ")"
            MsgBox "SharePointへのアップロードが完了しました。", vbInformation, "完了"
        Case -1
            ' 通信例外は SendHttpPost 内でログ済み
        Case Else
            MsgBox "アップロードに失敗しました。(HTTP " & httpStatus & ")", vbCritical, "エラー"
    End Select
End Sub

' ============================================================
' UploadAllToSharePoint — all シートの全データを Power Automate 経由でアップロード
'
' all シートの 2行目以降を Variant 配列に一括読み込みし、
' 全 11 フィールドを JSON オブジェクトに変換して送信する。
' ============================================================
Public Sub UploadAllToSharePoint()
    Dim paUrl As String
    Dim wsAll As Worksheet
    Dim lastRow As Long
    Dim allData As Variant
    Dim r As Long
    Dim rowsJson As String
    Dim sep As String
    Dim jsonBody As String
    Dim httpStatus As Long

    ' M3（全データ送信用）を優先し、未設定の場合は M2 にフォールバック
    paUrl = LoadPowerAutomateUrlAll()
    If paUrl = "" Then
        MsgBox "ConfigシートのM2またはM3にPowerAutomate URLが設定されていません。", _
               vbExclamation, "設定エラー"
        Exit Sub
    End If

    Set wsAll = ThisWorkbook.Sheets(SH_ALL)
    lastRow = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "allシートにデータがありません。先にファイルを読み込んでください。", _
               vbExclamation, "データなし"
        Exit Sub
    End If

    ' --- データ行を Variant 配列に一括読み込み（逐次セルアクセスを回避）---
    allData = wsAll.Range(wsAll.Cells(2, 1), wsAll.Cells(lastRow, ALL_TOTAL_COLS)).Value

    ' --- 各行を JSON オブジェクトに変換して連結 ---
    rowsJson = ""
    sep = ""
    For r = 1 To UBound(allData, 1)
        rowsJson = rowsJson & sep & "{"
        rowsJson = rowsJson & """client"":"    & JsonString(allData(r, ALL_COL_CLIENT))     & ","
        rowsJson = rowsJson & """prodCode"":"  & JsonString(allData(r, ALL_COL_PROD_CODE))  & ","
        rowsJson = rowsJson & """amount"":"    & JsonNumber(allData(r, ALL_COL_AMOUNT))     & ","
        rowsJson = rowsJson & """unitPrice"":" & JsonNumber(allData(r, ALL_COL_UNIT_PRICE)) & ","
        rowsJson = rowsJson & """qty"":"       & JsonNumber(allData(r, ALL_COL_QTY))        & ","
        rowsJson = rowsJson & """date"":"      & JsonString(allData(r, ALL_COL_DATE))       & ","
        rowsJson = rowsJson & """saleType"":"  & JsonString(allData(r, ALL_COL_SALE_TYPE))  & ","
        rowsJson = rowsJson & """dept"":"      & JsonString(allData(r, ALL_COL_DEPT))       & ","
        rowsJson = rowsJson & """prodName"":"  & JsonString(allData(r, ALL_COL_PROD_NAME))  & ","
        rowsJson = rowsJson & """margin"":"    & JsonNumber(allData(r, ALL_COL_MARGIN))     & ","
        rowsJson = rowsJson & """source"":"    & JsonString(allData(r, ALL_COL_SOURCE))
        rowsJson = rowsJson & "}"
        sep = ","
    Next r

    ' --- JSON ペイロード組み立て ---
    jsonBody = "{"
    jsonBody = jsonBody & """uploadedAt"":" & JsonString(Format(Now(), "yyyy/mm/dd hh:mm:ss")) & ","
    jsonBody = jsonBody & """rows"":["      & rowsJson & "]"
    jsonBody = jsonBody & "}"

    ' --- HTTP POST 送信と結果処理 ---
    httpStatus = SendHttpPost(paUrl, jsonBody)
    Select Case httpStatus
        Case 200, 202
            LogMessage "allシートSharePointアップロード完了 (HTTP " & httpStatus & "): " & (lastRow - 1) & "行"
            MsgBox "allシートのSharePointへのアップロードが完了しました。" & vbCrLf & _
                   (lastRow - 1) & "件のデータを送信しました。", vbInformation, "完了"
        Case -1
            ' 通信例外は SendHttpPost 内でログ済み
        Case Else
            MsgBox "アップロードに失敗しました。(HTTP " & httpStatus & ")", vbCritical, "エラー"
    End Select
End Sub

' ============================================================
' SendHttpPost — HTTP POST を送信して HTTP ステータスコードを返す（プライベート）
'
' 引数:
'   url      — 送信先 URL (Power Automate HTTP 要求トリガー)
'   jsonBody — 送信する JSON 文字列
'
' 戻り値:
'   HTTP ステータスコード (Long)
'     200 / 202 : 正常終了 (Power Automate は 202 Accepted を返すことが多い)
'     その他正数 : HTTP エラー (呼び出し元で MsgBox 表示)
'     -1         : 通信例外（このサブ内でログに記録済み）
'
' 設計方針:
'   メッセージの表示は呼び出し元が担当する。
'   このサブは通信のみを責務とし、例外発生時のみログを書く。
' ============================================================
Private Function SendHttpPost(url As String, jsonBody As String) As Long
    On Error GoTo HttpErr

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False           ' False = 同期通信（完了まで待機）
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonBody

    SendHttpPost = http.Status
    Exit Function

HttpErr:
    LogMessage "[エラー] HTTP通信例外: " & Err.Description
    SendHttpPost = -1
End Function

' ============================================================
' JsonString — Variant 値を JSON 文字列リテラルに変換する（プライベート）
'
' 戻り値: ダブルクォートで囲まれた JSON 文字列（特殊文字はエスケープ済み）
'
' エスケープ対象:  \ → \\  / " → \"  CR → \r  LF → \n  Tab → \t
' ※バックスラッシュを最初に処理することで二重エスケープを防ぐ
' ============================================================
Private Function JsonString(s As Variant) As String
    Dim str As String
    str = CStr(s)
    str = Replace(str, "\",   "\\")
    str = Replace(str, """",  "\""")
    str = Replace(str, Chr(13), "\r")
    str = Replace(str, Chr(10), "\n")
    str = Replace(str, Chr(9),  "\t")
    JsonString = """" & str & """"
End Function

' ============================================================
' JsonNumber — Variant 値を JSON 数値リテラルに変換する（プライベート）
'
' 戻り値:
'   数値の場合 : JSON 数値文字列（整数は小数点なし、小数はピリオド表記）
'   非数値の場合: "0"
'
' ロケール対応:
'   CStr の小数点記号は実行環境のロケールに依存するため、
'   Application.International(xlDecimalSeparator) でピリオンに統一する。
' ============================================================
Private Function JsonNumber(v As Variant) As String
    If Not IsNumeric(v) Then
        JsonNumber = "0"
        Exit Function
    End If
    Dim n As Double
    n = CDbl(v)
    If n = Int(n) Then
        JsonNumber = CStr(CLng(n))   ' 整数: 小数点なし
    Else
        JsonNumber = Replace(CStr(n), _
                             Application.International(xlDecimalSeparator), ".")
    End If
End Function
