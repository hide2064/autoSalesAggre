Option Explicit

' ============================================================
' autoSalesAggre workbook setup script
' Usage: cscript setup\create_workbook.vbs
' Prereq: Enable "Trust access to the VBA project object model"
'         in Excel Trust Center settings before running.
' ============================================================

Dim fso, scriptDir, srcPath, outputFile
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
srcPath   = fso.GetAbsolutePathName(fso.BuildPath(scriptDir, "..\src")) & "\"
outputFile = fso.GetAbsolutePathName(fso.BuildPath(scriptDir, "..\autoSalesAggre.xlsm"))

WScript.Echo "Setup started"
WScript.Echo "Source : " & srcPath
WScript.Echo "Output : " & outputFile

Dim xlApp, wb
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Set wb = xlApp.Workbooks.Add

' ---- Create 4 sheets: main, Config, all, Shuukei (placeholder for 集計) ----
' Remove extra default sheets first
Do While wb.Sheets.Count > 1
    wb.Sheets(wb.Sheets.Count).Delete
Loop
wb.Sheets(1).Name = "main"
Dim shConfig, shAll, shShuukei
' Add sheets without named args; they land before index 1, then rename in order
Set shConfig  = wb.Sheets.Add()
shConfig.Name = "Config"
Set shAll     = wb.Sheets.Add()
shAll.Name    = "all"
Set shShuukei = wb.Sheets.Add()
shShuukei.Name = "Shuukei"
' Arrange order: main(1), Config(2), all(3), Shuukei(4)
wb.Sheets("main").Move    wb.Sheets(1)
wb.Sheets("Config").Move  wb.Sheets(2)
wb.Sheets("all").Move     wb.Sheets(3)
wb.Sheets("Shuukei").Move wb.Sheets(4)

' ---- Import VBA modules ----
WScript.Echo "Importing VBA modules..."
Dim vbp
Set vbp = wb.VBProject

vbp.VBComponents.Import srcPath & "modConfig.bas"
vbp.VBComponents.Import srcPath & "modFileIO.bas"
vbp.VBComponents.Import srcPath & "modDataProcess.bas"
vbp.VBComponents.Import srcPath & "modAggregation.bas"
vbp.VBComponents.Import srcPath & "modUIControl.bas"
vbp.VBComponents.Import srcPath & "modSharePoint.bas"
vbp.VBComponents.Import srcPath & "modSetup.bas"
WScript.Echo "Modules imported."

' ---- Run one-time setup (renames sheets, sets layouts, injects event) ----
WScript.Echo "Running InitWorkbook..."
xlApp.Run "modSetup.InitWorkbook"
WScript.Echo "InitWorkbook complete."

' ---- Save as .xlsm (52 = xlOpenXMLMacroEnabled) ----
wb.SaveAs outputFile, 52
WScript.Echo "Saved: " & outputFile

xlApp.Quit
Set xlApp = Nothing

WScript.Echo ""
WScript.Echo "Done. Open autoSalesAggre.xlsm in Excel and enable macros."
