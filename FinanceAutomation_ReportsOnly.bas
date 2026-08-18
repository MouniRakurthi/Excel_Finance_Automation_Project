Attribute VB_Name = "FinanceAutomation"
'===================================================================================
'  FINANCE OPERATIONS AUTOMATION - VBA MODULE (Reports-only version)
'  Workbook:  Finance_Automation_Workbook.xlsm
'
'  HOW TO INSTALL:
'   1. Open Finance_Automation_Workbook.xlsx in desktop Excel.
'   2. File > Save As > choose "Excel Macro-Enabled Workbook (*.xlsm)".
'   3. Save it to a real local folder (not OneDrive) - e.g. C:\FinanceProject\
'   4. Press Alt+F11 to open the VBA Editor.
'   5. Insert a new Module, paste this entire file into it.
'   6. Press F5 on RunAllReports once to test, or add a worksheet button
'      (Developer tab > Insert > Button (Form Control)) and assign it to
'      RunAllReports for a one-click refresh + export.
'
'  WHAT EACH MACRO DOES
'   RefreshAllReports  - recalculates every formula and refreshes any
'                         PivotTables you have built on top of this data.
'   ExportReportsToPDF - exports the four recurring report tabs plus the
'                         KPI Dashboard as a single dated PDF snapshot.
'   RunAllReports       - runs both above in sequence. This is the
'                         one-click button macro.
'
'  SCHEDULING (optional, VBA cannot run while Excel is closed)
'   For "every Monday at 9am" automatic report generation, pair this
'   workbook with Windows Task Scheduler:
'     - Action: Start a program
'     - Program/script:  full path to EXCEL.EXE
'     - Add arguments:   "full path to this workbook" /x
'     - Rename Auto_Open_Template (below) to Auto_Open and uncomment its
'       line so the macro fires automatically once the file opens.
'   The How-To-Run Guide walks through the exact Task Scheduler screens.
'===================================================================================

Option Explicit

'----------------------- SETTINGS (edit these) -----------------------------------
Const REPORT_SHEETS As String = "Report_Weekly,Report_Monthly,Report_Quarterly,Report_Yearly,KPI_Dashboard"
Const EXPORT_FOLDER As String = ""   ' leave blank to export next to the workbook
'-----------------------------------------------------------------------------------


Public Function GetExportFolder() As String
    If Len(EXPORT_FOLDER) > 0 Then
        GetExportFolder = EXPORT_FOLDER
    Else
        GetExportFolder = ThisWorkbook.Path & "\Reports_Exported\"
    End If
    If Dir(GetExportFolder, vbDirectory) = "" Then
        MkDir GetExportFolder
    End If
End Function


'===================================================================================
' 1) REFRESH - recalculate everything and refresh any PivotTables/PivotCaches
'===================================================================================
Sub RefreshAllReports()
    Dim ws As Worksheet
    Dim pt As PivotTable

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild        ' forces every formula, incl. volatile ones, to re-run

    ' Refresh any PivotTables the user has added on top of the module tabs
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    Application.ScreenUpdating = True
End Sub


'===================================================================================
' 2) EXPORT - save the recurring report tabs + dashboard as one PDF snapshot
'===================================================================================
Function ExportReportsToPDF() As String
    Dim sheetNames() As String
    Dim i As Long
    Dim exportPath As String

    sheetNames = Split(REPORT_SHEETS, ",")

    ' Select all report sheets together so they export as one multi-page PDF
    ThisWorkbook.Sheets(sheetNames(0)).Select
    For i = 1 To UBound(sheetNames)
        ThisWorkbook.Sheets(sheetNames(i)).Select False
    Next i

    exportPath = GetExportFolder() & "Finance_Ops_Report_" & Format(Now, "yyyy-mm-dd_hhnn") & ".pdf"

    ActiveSheet.Range("A1").Select    ' avoid exporting a stray selection
    ThisWorkbook.ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=exportPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ThisWorkbook.Sheets(sheetNames(0)).Select   ' clear the multi-sheet selection
    ExportReportsToPDF = exportPath
End Function


'===================================================================================
' 3) MASTER MACRO - refresh + export in one click
'===================================================================================
Sub RunAllReports()
    Dim pdfPath As String

    RefreshAllReports
    pdfPath = ExportReportsToPDF()

    MsgBox "Reports refreshed and exported:" & vbCrLf & vbCrLf & pdfPath, vbInformation, "Finance Reports Ready"
End Sub


'===================================================================================
' OPTIONAL - fire automatically when the file is opened by Task Scheduler
' Rename this to Auto_Open (exact name) if you want it to run on file-open.
'===================================================================================
Sub Auto_Open_Template()
    ' Rename to "Auto_Open" to activate. Left disabled by default so opening the
    ' file manually during setup/testing does not immediately export a PDF.
    ' RunAllReports
End Sub
