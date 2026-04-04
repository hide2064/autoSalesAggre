Attribute VB_Name = "modSharePoint"
Option Explicit

' SharePoint upload via Power Automate HTTP trigger
' Config sheet M2 に Power Automate の HTTP POST URL を設定してください。
'
' Power Automate フローは以下の JSON を受け取ります:
'   {
'     "dept":       "全部署",
'     "fromDate":   "2026/01/01",
'     "toDate":     "2026/03/31",
'     "uploadedAt": "2026/04/04 12:00:00",
'     "rows": [
'       {"name": "製品A",     "amount": 100000, "qty": 10, "margin": 10000},
'       {"name": "  客先X",   "amount":  80000, "qty":  8, "margin":  8000},
'       ...
'       {"name": "総合計",    "amount": 200000, "qty": 20, "margin": 20000}
'     ]
'   }

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
    Dim rowsJson As String
    Dim sep As String
    Dim jsonBody As String
    Dim http As Object

    paUrl = LoadPowerAutomateUrl()
    If paUrl = "" Then
        MsgBox "ConfigシートのM2にPowerAutomate URLが設定されていません。", _
               vbExclamation, "設定エラー"
        Exit Sub
    End If

    Set wsAggr = ThisWorkbook.Sheets(SH_AGGR)

    dept     = Trim(CStr(wsAggr.Range(AGGR_DEPT_CELL).Value))
    fromDate = Trim(CStr(wsAggr.Range(AGGR_FROM_CELL).Value))
    toDate   = Trim(CStr(wsAggr.Range(AGGR_TO_CELL).Value))

    lastRow = wsAggr.Cells(wsAggr.Rows.Count, 1).End(xlUp).Row
    If lastRow < AGGR_DATA_ROW Then
        MsgBox "集計データがありません。先にデータを集計してください。", _
               vbExclamation, "データなし"
        Exit Sub
    End If

    ' 集計テーブルの各行を JSON 配列に変換
    rowsJson = ""
    sep = ""
    For r = AGGR_DATA_ROW To lastRow
        itemName = Trim(CStr(wsAggr.Cells(r, 1).Value))
        If itemName = "" Then GoTo NextDataRow
        amt    = wsAggr.Cells(r, 2).Value
        qty    = wsAggr.Cells(r, 3).Value
        margin = wsAggr.Cells(r, 4).Value

        rowsJson = rowsJson & sep & "{"
        rowsJson = rowsJson & """name"":"   & JsonString(wsAggr.Cells(r, 1).Value) & ","
        rowsJson = rowsJson & """amount"":"  & JsonNumber(amt) & ","
        rowsJson = rowsJson & """qty"":"     & JsonNumber(qty) & ","
        rowsJson = rowsJson & """margin"":"  & JsonNumber(margin)
        rowsJson = rowsJson & "}"
        sep = ","
NextDataRow:
    Next r

    ' ペイロード組み立て
    jsonBody = "{"
    jsonBody = jsonBody & """dept"":"       & JsonString(dept)     & ","
    jsonBody = jsonBody & """fromDate"":"   & JsonString(fromDate)  & ","
    jsonBody = jsonBody & """toDate"":"     & JsonString(toDate)    & ","
    jsonBody = jsonBody & """uploadedAt"":" & JsonString(Format(Now(), "yyyy/mm/dd hh:mm:ss")) & ","
    jsonBody = jsonBody & """rows"":["      & rowsJson & "]"
    jsonBody = jsonBody & "}"

    ' HTTP POST 送信
    On Error GoTo HttpErr
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", paUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonBody

    If http.Status = 200 Or http.Status = 202 Then
        LogMessage "SharePointアップロード完了 (HTTP " & http.Status & ")"
        MsgBox "SharePointへのアップロードが完了しました。", vbInformation, "完了"
    Else
        LogMessage "[エラー] SharePointアップロード失敗 (HTTP " & http.Status & "): " & http.responseText
        MsgBox "アップロードに失敗しました。" & vbCrLf & _
               "HTTP " & http.Status & vbCrLf & http.responseText, vbCritical, "エラー"
    End If
    Exit Sub

HttpErr:
    LogMessage "[エラー] SharePointアップロード例外: " & Err.Description
    MsgBox "アップロード中にエラーが発生しました:" & vbCrLf & Err.Description, _
           vbCritical, "エラー"
End Sub

' ---- ヘルパー ----

Private Function JsonString(s As Variant) As String
    Dim str As String
    str = CStr(s)
    str = Replace(str, "\",  "\\")
    str = Replace(str, """", "\""")
    str = Replace(str, Chr(13), "\r")
    str = Replace(str, Chr(10), "\n")
    str = Replace(str, Chr(9),  "\t")
    JsonString = """" & str & """"
End Function

Private Function JsonNumber(v As Variant) As String
    If Not IsNumeric(v) Then
        JsonNumber = "0"
        Exit Function
    End If
    Dim n As Double
    n = CDbl(v)
    If n = Int(n) Then
        JsonNumber = CStr(CLng(n))
    Else
        ' ロケール非依存で小数点をピリオドに統一
        JsonNumber = Replace(CStr(n), _
                             Application.International(xlDecimalSeparator), ".")
    End If
End Function
