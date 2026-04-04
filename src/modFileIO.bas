Attribute VB_Name = "modFileIO"
Option Explicit

Public Function SelectFiles() As Variant
    ' Returns Variant array of paths, or Boolean False if cancelled
    SelectFiles = Application.GetOpenFilename( _
        FileFilter:="テキストファイル (*.txt;*.tsv),*.txt;*.tsv", _
        Title:="読み込むTSVファイルを選択してください", _
        MultiSelect:=True)
End Function

Public Function LoadTsvToSheet(filePath As String) As Boolean
    ' Creates or replaces a sheet named after the file. Returns True on success.
    On Error GoTo ErrHandler

    Dim sheetName As String
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim fileNum As Integer
    Dim rowNum As Long
    Dim lineText As String
    Dim cols() As String
    Dim c As Integer

    sheetName = FilePathToSheetName(filePath)

    ' Delete existing sheet with same name
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next ws

    ' Insert before 集計 sheet to keep sheet order consistent
    Set newSheet = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(SH_AGGR))
    newSheet.Name = sheetName

    ' Read TSV line by line, store all values as text to preserve leading zeros
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    rowNum = 1
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText

        cols = Split(lineText, vbTab)

        For c = 0 To UBound(cols)
            With newSheet.Cells(rowNum, c + 1)
                .NumberFormat = "@"
                .Value = cols(c)
            End With
        Next c
        rowNum = rowNum + 1
    Loop

    Close #fileNum
    LoadTsvToSheet = True
    Exit Function

ErrHandler:
    Application.DisplayAlerts = True
    If fileNum > 0 Then Close #fileNum
    LoadTsvToSheet = False
End Function

Private Function FilePathToSheetName(filePath As String) As String
    ' Extract filename without extension; strip invalid sheet name chars; truncate to 31 chars
    Dim fileName As String
    Dim dotPos As Integer
    Dim invalids As String
    Dim i As Integer

    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then fileName = Left(fileName, dotPos - 1)

    ' Remove characters invalid for sheet names: \ / ? * [ ] :
    invalids = "\/?*[]:"
    For i = 1 To Len(invalids)
        fileName = Join(Split(fileName, Mid(invalids, i, 1)), "_")
    Next i

    If Len(fileName) > 31 Then fileName = Left(fileName, 31)
    FilePathToSheetName = fileName
End Function
