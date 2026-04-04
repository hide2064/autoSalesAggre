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
'   3. MSXML2.XMLHTTP を使って HTTP POST で送信
'   4. HTTP ステータス 200/202 を正常終了とみなす
'
' Power Automate フローの設定:
'   ・トリガー: "HTTP 要求を受信したとき" (When a HTTP request is received)
'   ・Content-Type: application/json
'   ・URL を Config シートの M2 に貼り付ける
'
' 送信 JSON フォーマット（集計シート版）:
'   {
'     "dept":       "全部署",          // 部署フィルタ値
'     "fromDate":   "2026/01/01",      // 開始日フィルタ値（空欄時は ""）
'     "toDate":     "2026/03/31",      // 終了日フィルタ値（空欄時は ""）
'     "uploadedAt": "2026/04/04 12:00:00",
'     "rows": [
'       {"name": "製品A",       "amount": 150000, "qty": 30, "margin": 15000},
'       {"name": "　　得意先X", "amount": 100000, "qty": 20, "margin": 10000},
'       ...
'       {"name": "総合計",      "amount": 200000, "qty": 40, "margin": 20000}
'     ]
'   }
'
' 送信 JSON フォーマット（all シート版）:
'   {
'     "uploadedAt": "2026/04/04 12:00:00",
'     "rows": [
'       {
'         "client":   "客先A",    "prodCode": "P001",
'         "amount":   10000,      "unitPrice": 1000,
'         "qty":      10,         "date":      "2026/01/15",
'         "saleType": "直販",     "dept":      "営業部",
'         "prodName": "製品A",    "margin":    1000,
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
    Dim itemName As String
    Dim amt As Variant
    Dim qty As Variant
    Dim margin As Variant
    Dim dept As String
    Dim fromDate As String
    Dim toDate As String
    Dim rowsJson As String  ' rows 配列の JSON 文字列（累積）
    Dim sep As String       ' 配列要素間のカンマ（初回は空）
    Dim jsonBody As String
    Dim http As Object

    ' --- URL の取得と未設定チェック ---
    paUrl = LoadPowerAutomateUrl()
    If paUrl = "" Then
        MsgBox "ConfigシートのM2にPowerAutomate URLが設定されていません。", _
               vbExclamation, "設定エラー"
        Exit Sub
    End If

    Set wsAggr = ThisWorkbook.Sheets(SH_AGGR)

    ' --- フィルタ条件の読み取り（JSON の上位フィールドに含める）---
    dept     = Trim(CStr(wsAggr.Range(AGGR_DEPT_CELL).Value))
    fromDate = Trim(CStr(wsAggr.Range(AGGR_FROM_CELL).Value))
    toDate   = Trim(CStr(wsAggr.Range(AGGR_TO_CELL).Value))

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
        itemName = Trim(CStr(wsAggr.Cells(r, 1).Value))
        If itemName = "" Then GoTo NextDataRow  ' 空行はスキップ

        amt    = wsAggr.Cells(r, 2).Value
        qty    = wsAggr.Cells(r, 3).Value
        margin = wsAggr.Cells(r, 4).Value

        ' name には字下げ込みの表示文字列（"　　客先名" や "総合計"）をそのまま使用
        rowsJson = rowsJson & sep & "{"
        rowsJson = rowsJson & """name"":"   & JsonString(wsAggr.Cells(r, 1).Value) & ","
        rowsJson = rowsJson & """amount"":"  & JsonNumber(amt) & ","
        rowsJson = rowsJson & """qty"":"     & JsonNumber(qty) & ","
        rowsJson = rowsJson & """margin"":"  & JsonNumber(margin)
        rowsJson = rowsJson & "}"
        sep = ","
NextDataRow:
    Next r

    ' --- JSON ペイロード組み立て ---
    jsonBody = "{"
    jsonBody = jsonBody & """dept"":"       & JsonString(dept)     & ","
    jsonBody = jsonBody & """fromDate"":"   & JsonString(fromDate)  & ","
    jsonBody = jsonBody & """toDate"":"     & JsonString(toDate)    & ","
    jsonBody = jsonBody & """uploadedAt"":" & JsonString(Format(Now(), "yyyy/mm/dd hh:mm:ss")) & ","
    jsonBody = jsonBody & """rows"":["      & rowsJson & "]"
    jsonBody = jsonBody & "}"

    ' --- HTTP POST 送信 ---
    Call SendHttpPost(paUrl, jsonBody, "SharePointアップロード完了", "SharePointアップロード失敗", "SharePointアップロード例外")
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
    Dim allData As Variant  ' all シートデータの一括読み込み用
    Dim r As Long
    Dim rowsJson As String
    Dim sep As String
    Dim jsonBody As String

    ' --- URL の取得と未設定チェック ---
    paUrl = LoadPowerAutomateUrl()
    If paUrl = "" Then
        MsgBox "ConfigシートのM2にPowerAutomate URLが設定されていません。", _
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

    ' --- HTTP POST 送信（成功時はログと件数メッセージを表示）---
    Dim http As Object
    On Error GoTo HttpErr
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", paUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonBody

    If http.Status = 200 Or http.Status = 202 Then
        LogMessage "allシートSharePointアップロード完了 (HTTP " & http.Status & "): " & (lastRow - 1) & "行"
        MsgBox "allシートのSharePointへのアップロードが完了しました。" & vbCrLf & _
               (lastRow - 1) & "件のデータを送信しました。", vbInformation, "完了"
    Else
        LogMessage "[エラー] allシートSharePointアップロード失敗 (HTTP " & http.Status & "): " & http.responseText
        MsgBox "アップロードに失敗しました。" & vbCrLf & _
               "HTTP " & http.Status & vbCrLf & http.responseText, vbCritical, "エラー"
    End If
    Exit Sub

HttpErr:
    LogMessage "[エラー] allシートSharePointアップロード例外: " & Err.Description
    MsgBox "アップロード中にエラーが発生しました:" & vbCrLf & Err.Description, _
           vbCritical, "エラー"
End Sub

' ============================================================
' SendHttpPost — HTTP POST 送信の共通処理（プライベート）
'
' 引数:
'   url          — 送信先 URL
'   jsonBody     — 送信する JSON 文字列
'   successLabel — ログに記録する成功時のラベル文字列
'   failLabel    — ログに記録する失敗時のラベル文字列
'   errLabel     — ログに記録する例外時のラベル文字列
'
' HTTP 200 または 202 を正常終了とみなす。
' Power Automate の HTTP 要求トリガーは受信後すぐに 202 Accepted を返す。
' ============================================================
Private Sub SendHttpPost(url As String, jsonBody As String, _
                         successLabel As String, failLabel As String, errLabel As String)
    Dim http As Object
    On Error GoTo HttpErr

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False           ' 第3引数 False = 同期通信（完了まで待機）
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonBody

    If http.Status = 200 Or http.Status = 202 Then
        LogMessage successLabel & " (HTTP " & http.Status & ")"
        MsgBox "SharePointへのアップロードが完了しました。", vbInformation, "完了"
    Else
        LogMessage "[エラー] " & failLabel & " (HTTP " & http.Status & "): " & http.responseText
        MsgBox "アップロードに失敗しました。" & vbCrLf & _
               "HTTP " & http.Status & vbCrLf & http.responseText, vbCritical, "エラー"
    End If
    Exit Sub

HttpErr:
    LogMessage "[エラー] " & errLabel & ": " & Err.Description
    MsgBox "アップロード中にエラーが発生しました:" & vbCrLf & Err.Description, _
           vbCritical, "エラー"
End Sub

' ============================================================
' JsonString — Variant 値を JSON 文字列リテラルに変換する（プライベート）
'
' 引数:
'   s — 変換する値（Variant; CStr で文字列化してから処理）
'
' 戻り値: ダブルクォートで囲まれた JSON 文字列
'   例: "製品A" → """製品A"""
'
' エスケープ処理:
'   バックスラッシュ → "\\"
'   ダブルクォート   → "\""
'   CR               → "\r"
'   LF               → "\n"
'   タブ             → "\t"
' ============================================================
Private Function JsonString(s As Variant) As String
    Dim str As String
    str = CStr(s)
    str = Replace(str, "\",   "\\")   ' バックスラッシュは最初にエスケープ（二重置換防止）
    str = Replace(str, """",  "\""")
    str = Replace(str, Chr(13), "\r")
    str = Replace(str, Chr(10), "\n")
    str = Replace(str, Chr(9),  "\t")
    JsonString = """" & str & """"
End Function

' ============================================================
' JsonNumber — Variant 値を JSON 数値リテラルに変換する（プライベート）
'
' 引数:
'   v — 変換する値（IsNumeric が False の場合は "0" を返す）
'
' 戻り値: JSON 数値文字列（整数は CLng で、小数はロケール非依存でピリオド表記）
'
' ロケール対応:
'   VBA の CStr は実行環境のロケールに依存するため、
'   小数点区切り文字がカンマになる環境（ヨーロッパ等）では
'   Application.International(xlDecimalSeparator) を使ってピリオドに変換する。
' ============================================================
Private Function JsonNumber(v As Variant) As String
    If Not IsNumeric(v) Then
        JsonNumber = "0"
        Exit Function
    End If
    Dim n As Double
    n = CDbl(v)
    If n = Int(n) Then
        ' 整数の場合は小数点なしで出力
        JsonNumber = CStr(CLng(n))
    Else
        ' 小数の場合はロケール固有の小数点記号をピリオンに統一
        JsonNumber = Replace(CStr(n), _
                             Application.International(xlDecimalSeparator), ".")
    End If
End Function
