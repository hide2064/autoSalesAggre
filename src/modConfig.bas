Attribute VB_Name = "modConfig"
Option Explicit

' ===== Config sheet table positions =====
Public Const CFG_PRODUCT_HDR_ROW    As Integer = 2   ' 製品マスタ header row (A2)
Public Const CFG_PRODUCT_COL        As Integer = 1   ' A: 製品コード
Public Const CFG_COMMISSION_HDR_ROW As Integer = 2   ' 口銭マスタ header row (D2)
Public Const CFG_COMMISSION_COL     As Integer = 4   ' D: 売上種別
Public Const CFG_HEADER_HDR_ROW     As Integer = 2   ' 名寄せ header row (G2)
Public Const CFG_HEADER_COL         As Integer = 7   ' G: 正規名
Public Const CFG_DEPT_HDR_ROW       As Integer = 2   ' 部署リスト header row (J2)
Public Const CFG_DEPT_COL           As Integer = 10  ' J: 部署リスト

' ===== all sheet column indices (1-based) =====
Public Const ALL_COL_CLIENT     As Integer = 1   ' 客先名
Public Const ALL_COL_PROD_CODE  As Integer = 2   ' 製品コード
Public Const ALL_COL_AMOUNT     As Integer = 3   ' 売上金額
Public Const ALL_COL_UNIT_PRICE As Integer = 4   ' 製品単価
Public Const ALL_COL_QTY        As Integer = 5   ' 売上数量
Public Const ALL_COL_DATE       As Integer = 6   ' 売上発生日
Public Const ALL_COL_SALE_TYPE  As Integer = 7   ' 売上種別
Public Const ALL_COL_DEPT       As Integer = 8   ' 部署
Public Const ALL_COL_PROD_NAME  As Integer = 9   ' 製品名 (calculated)
Public Const ALL_COL_MARGIN     As Integer = 10  ' 部署取り分 (calculated)
Public Const ALL_COL_SOURCE     As Integer = 11  ' ソースファイル
Public Const ALL_TOTAL_COLS     As Integer = 11

' ===== Sheet names =====
Public Const SH_MAIN   As String = "main"
Public Const SH_CONFIG As String = "Config"
Public Const SH_ALL    As String = "all"
Public Const SH_AGGR   As String = "集計"

' ===== 集計 sheet cell addresses =====
Public Const AGGR_DEPT_CELL As String = "B1"
Public Const AGGR_FROM_CELL As String = "B2"
Public Const AGGR_TO_CELL   As String = "B3"
Public Const AGGR_HDR_ROW   As Integer = 5
Public Const AGGR_DATA_ROW  As Integer = 6

' ===== main sheet =====
Public Const MAIN_LOG_START_ROW As Integer = 3

' ---------- Master loading ----------

Public Function LoadProductDict() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_CONFIG)

    Dim r As Long
    r = CFG_PRODUCT_HDR_ROW + 1
    Do While Trim(CStr(ws.Cells(r, CFG_PRODUCT_COL).Value)) <> ""
        Dim code As String
        code = Trim(CStr(ws.Cells(r, CFG_PRODUCT_COL).Value))
        If Not dict.Exists(code) Then
            dict(code) = Trim(CStr(ws.Cells(r, CFG_PRODUCT_COL + 1).Value))
        End If
        r = r + 1
    Loop

    Set LoadProductDict = dict
End Function

Public Function LoadCommissionDict() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_CONFIG)

    Dim r As Long
    r = CFG_COMMISSION_HDR_ROW + 1
    Do While Trim(CStr(ws.Cells(r, CFG_COMMISSION_COL).Value)) <> ""
        Dim stype As String
        stype = Trim(CStr(ws.Cells(r, CFG_COMMISSION_COL).Value))
        If Not dict.Exists(stype) Then
            dict(stype) = CDbl(ws.Cells(r, CFG_COMMISSION_COL + 1).Value)
        End If
        r = r + 1
    Loop

    Set LoadCommissionDict = dict
End Function

Public Function LoadHeaderMap() As Object
    ' Returns dict: LCase(trimmed_alias) -> canonical_column_name
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_CONFIG)

    Dim r As Long
    r = CFG_HEADER_HDR_ROW + 1
    Do While Trim(CStr(ws.Cells(r, CFG_HEADER_COL).Value)) <> ""
        Dim canonical As String
        canonical = Trim(CStr(ws.Cells(r, CFG_HEADER_COL).Value))
        Dim aliases As String
        aliases = Trim(CStr(ws.Cells(r, CFG_HEADER_COL + 1).Value))

        ' Register canonical name itself
        If Not dict.Exists(LCase(canonical)) Then dict(LCase(canonical)) = canonical

        ' Register each alias
        Dim parts() As String
        parts = Split(aliases, ",")
        Dim i As Integer
        For i = 0 To UBound(parts)
            Dim a As String
            a = LCase(Trim(parts(i)))
            If a <> "" And Not dict.Exists(a) Then dict(a) = canonical
        Next i
        r = r + 1
    Loop

    Set LoadHeaderMap = dict
End Function

Public Sub RefreshDeptList(dictDept As Object)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_CONFIG)

    ' Clear J3 downward
    Dim clearRow As Long
    clearRow = CFG_DEPT_HDR_ROW + 1
    Do While Trim(CStr(ws.Cells(clearRow, CFG_DEPT_COL).Value)) <> ""
        ws.Cells(clearRow, CFG_DEPT_COL).ClearContents
        clearRow = clearRow + 1
    Loop

    ' J2 = "全部署" (fixed)
    ws.Cells(CFG_DEPT_HDR_ROW, CFG_DEPT_COL).Value = "全部署"

    ' Write unique depts from J3
    Dim r As Long
    r = CFG_DEPT_HDR_ROW + 1
    Dim key As Variant
    For Each key In dictDept.Keys
        ws.Cells(r, CFG_DEPT_COL).Value = key
        r = r + 1
    Next key

    Dim lastDeptRow As Long
    lastDeptRow = r - 1

    ' Update 集計!B1 dropdown
    Dim wsAggr As Worksheet
    Set wsAggr = ThisWorkbook.Sheets(SH_AGGR)
    With wsAggr.Range(AGGR_DEPT_CELL).Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=" & SH_CONFIG & "!$J$" & CFG_DEPT_HDR_ROW & ":$J$" & lastDeptRow
    End With

    If Trim(CStr(wsAggr.Range(AGGR_DEPT_CELL).Value)) = "" Then
        wsAggr.Range(AGGR_DEPT_CELL).Value = "全部署"
    End If
End Sub
