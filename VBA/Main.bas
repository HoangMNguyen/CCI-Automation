Attribute VB_Name = "Main"
Option Explicit

Sub OutFormat() 'Format, sort column A ascending,
showAll
ActiveSheet.Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Font.Bold = True
Range(Selection, Selection.End(xlDown)).Select
Selection.WrapText = False
Selection.Columns.AutoFit
Selection.VerticalAlignment = xlCenter
Selection.HorizontalAlignment = xlLeft
Selection.Borders.LineStyle = xlContinuous
Selection.Borders.Weight = xlThin
If Not ActiveSheet.AutoFilterMode Then
    Selection.AutoFilter
End If
Selection.Font.Size = 10
Selection.Borders.LineStyle = xlContinuous
Selection.Borders.Color = vbBlack
Selection.Borders.Weight = xlThin

sortAsc ("A1")
End Sub


Sub QuickReportsFormStatusFormatSponsor() 'reformat quick report for form status report, split into cohorts depend on studies
' Keyboard Shortcut: Ctrl+Shift+F
    Dim lastRow As Long
    Dim WS1 As Worksheet
    Dim WS2 As Worksheet
    Dim WS3 As Worksheet
    Dim WS4 As Worksheet
    Dim SheetNum As Long
    Dim VSheet As Worksheet
    Dim TempSheet As Worksheet
    Dim xcell As Object
    Dim SelectCells As Range
    Dim fileSaveName As Variant
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set WS1 = Sheets(1)
    WS1.Activate
    WS1.Name = "All Cohorts Form Status Report"
    
    'Counting LastRow and LastCol for studies which only have 1 output, Errl will skip over the code above and continue loop
    WS1.Range("A1").Select
    lastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
    'Delete last row
    Range("A" & lastRow).Select
    Selection.EntireRow.Delete
    lastRow = lastRow - 1
   
   'Delete More Subject Details forms
    If Application.WorksheetFunction.CountIf(WS1.Range("J2:O" & lastRow), "More Subject Details") > 0 Then
        ActiveSheet.Range("A1").AutoFilter Field:=10, Criteria1:="More Subject Details"
        WS1.Range("A2:A" & lastRow).Select
        Selection.EntireRow.Delete
    End If
    Range("A:B,D:D,F:F,I:I").Select
    Range("I1").Activate
    Selection.Delete Shift:=xlToLeft
    Rows("1:1").Select
    Selection.AutoFilter
    


    
    
    If WS1.Range("A2").Value = "827644" Then
        Call S14217.QFSR(WS1, WS2, WS3, WS4, lastRow)
        WS1.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"
        WS2.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"
        WS3.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"
        WS4.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"
    
    ElseIf WS1.Range("A2").Value = "826085" Then
        Call S02916.QFSR(WS1, WS2, WS3, WS4, lastRow)
        WS1.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"
        WS2.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"
        WS3.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"

    ElseIf WS1.Range("A2").Value = "850925" Then
        Call S01422.QFSR(WS1, "Submitted to Sponsor")


    ElseIf WS1.Range("A2").Value = "823312" Then
        Call S15CT055.QFSR(WS1, WS2, WS3, WS4, lastRow)
        WS1.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"
        WS2.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"
        WS3.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"
    

    ElseIf WS1.Range("A2").Value = "826250" Then
        Call S32816.QFSR(WS1, WS2, WS3, WS4, lastRow)
        WS1.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"
        WS2.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"
        WS3.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"
        WS4.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"
    
    
    Else 'other studies
        WS1.Name = "Form Status Report"
        Set WS2 = Sheets.Add
        WS2.Name = "Form Status Overview"
        WS2.Range("A1").Value = "Form Status"
        Call FormStatusOverview(WS1, WS2, 1, lastRow)
        
        'Autofit and add borders for the form status overview table
        ActiveSheet.Range("A1").Select
        Call FormatTable
        
        WS1.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"
    End If
    
    'add validation tab
    Sheets.Add After:=Sheets(Sheets.Count)
    Set VSheet = Sheets(Sheets.Count)
    VSheet.Name = "Validation tab"
    Sheets.Add After:=Sheets(Sheets.Count)
    Set TempSheet = Sheets(Sheets.Count)
    

    WS1.Activate
    Call RemoveFilter
    WS1.Range("A1", WS1.Range("J1").End(xlDown)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=TempSheet.Range("A1")

    TempSheet.Activate
    
    For Each xcell In ActiveSheet.Range("F1:F" & lastRow).Cells
    If xcell.Text = "Incomplete" Or xcell.Text = "Work In Progress" Or xcell.Text = "Submitted to Sponsor" Then
        If SelectCells Is Nothing Then
            Set SelectCells = Range(xcell.Address)
        Else
            Set SelectCells = Union(SelectCells, Range(xcell.Address))
        End If
    End If
    Next
    If Not SelectCells Is Nothing Then
        SelectCells.EntireRow.Delete
    End If
    Set SelectCells = Nothing
    
    TempSheet.Range("J1", TempSheet.Range("J1").End(xlDown)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=VSheet.Range("A1"), Unique:=True
    Dim TempRow As Long
    TempRow = TempSheet.Range("J1", TempSheet.Range("J1").End(xlDown)).Rows.Count
    
    
    VSheet.Activate
    VSheet.Range("A1").Value = "Completed/Ready for Submission Last Modified By"
    VSheet.Range("B1").Value = "Count of occurrence"
    VSheet.Range("A1").Select
    Call FormatTable
    
    Dim lngCount As Long
    
    
    lngCount = Application.WorksheetFunction.CountA(Columns(1))
    Dim cel As Range
    TempSheet.Activate
    For Each cel In VSheet.Range("A2:A" & lngCount).Cells
        VSheet.Range("B" & cel.Row).Value = Application.WorksheetFunction.CountIf(Range("J1:J" & TempRow), cel.Text)
        If cel.Value = "" Then
            cel.Value = "Detected form(s) created by site with Ready for Submission/Completed status"
            cel.Font.Bold = True
        End If
    Next cel
    
    Call RemoveFilter
    WS1.Range("A1").AutoFilter Field:=3, Criteria1:="="
    WS1.Range("A1", WS1.Range("J1").End(xlDown)).SpecialCells(xlCellTypeVisible).Copy
    VSheet.Range("D1").PasteSpecial
    VSheet.Activate
    VSheet.Range("D1").Select
    Call FormatTable
    
    TempSheet.Delete
    Application.ScreenUpdating = True
    'switching on the alert button
    
    Application.DisplayAlerts = True
    
    WS1.Activate
    Call RemoveFilter
    ActiveSheet.Range("A1").Select
    Call FormatTable
    WS1.Range("A1").AutoFilter Field:=6, Criteria1:="Submitted to Sponsor"
    
    Sheets(2).Activate
    
    'Save file
    Dim modifiedDate As String
    modifiedDate = Now2Date(Now)
    Dim modifiedTime As String
    modifiedTime = Now2Time(Now)
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:=modifiedDate & "-XXXXX Form Status Report " & modifiedTime & " EST_Sponsor " & ".xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    
    If fileSaveName = False Then
        MsgBox "You haven't saved the document", vbExclamation
    Else
        ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
    End If
    
End Sub

Sub QuickReportsFormStatusFormatSite() 'reformat quick report for form status report SITE VERSION, split into cohorts depend on studies
    
    'Keyboard Shortcut: Ctrl+Shift+N
    Dim NewBook As Workbook
    Dim xcell As Object
    Dim SelectCells As Range
    Dim lastRow As Long
    Dim WS1 As Worksheet
    Dim WS2 As Worksheet
    Dim WS3 As Worksheet
    Dim WS4 As Worksheet
    Dim SheetNum As Long
    Dim CurrentSheetNum As Long
    Dim fileSaveName As Variant
    Set NewBook = duplicateWorkbook(ActiveWorkbook)
    'Disable Screen Update
    Application.ScreenUpdating = False
    NewBook.Activate
    Set WS1 = Sheets(1)
    WS1.Activate
    WS1.Name = "All Cohorts Form Status Report"
    
    'Counting LastRow and LastCol for studies which only have 1 output, Errl will skip over the code above and continue loop
    WS1.Range("A1").Select
    lastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
    'Delete last row
    Range("A" & lastRow).Select
    Selection.EntireRow.Delete
    lastRow = lastRow - 1
   
   'Delete More Subject Details forms
    If Application.WorksheetFunction.CountIf(WS1.Range("J2:O" & lastRow), "More Subject Details") > 0 Then
        ActiveSheet.Range("A1").AutoFilter Field:=10, Criteria1:="More Subject Details"
        WS1.Range("A2:A" & lastRow).Select
        Selection.EntireRow.Delete
        
    End If
    Range("A:B,D:D,F:F,I:I").Select
    Range("I1").Activate
    Selection.Delete Shift:=xlToLeft
    Rows("1:1").Select
    Selection.AutoFilter
    
    If WS1.Range("A2").Value = "827644" Then
        Call S14217.QFSR(WS1, WS2, WS3, WS4, lastRow)
        WS1.Range("A1").AutoFilter Field:=6, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"
        WS2.Range("A1").AutoFilter Field:=6, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"
        WS3.Range("A1").AutoFilter Field:=6, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"
        WS4.Range("A1").AutoFilter Field:=6, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"
    
    ElseIf WS1.Range("A2").Value = "826085" Then
        Call S02916.QFSR(WS1, WS2, WS3, WS4, lastRow)
        WS1.Range("A1").AutoFilter Field:=6, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"
        WS2.Range("A1").AutoFilter Field:=6, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"
        WS3.Range("A1").AutoFilter Field:=6, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"
                
    ElseIf WS1.Range("A2").Value = "850925" Then
        Call S01422.QFSR(WS1, "Incomplete", "Work In Progress")

    ElseIf WS1.Range("A2").Value = "823312" Then
        Call S15CT055.QFSR(WS1, WS2, WS3, WS4, lastRow)
        WS1.Range("A1").AutoFilter Field:=6, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"
        WS2.Range("A1").AutoFilter Field:=6, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"
        WS3.Range("A1").AutoFilter Field:=6, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"


    ElseIf WS1.Range("A2").Value = "826250" Then
        Call S32816.QFSR(WS1, WS2, WS3, WS4, lastRow)
        WS1.Range("A1").AutoFilter Field:=6, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"
        WS2.Range("A1").AutoFilter Field:=6, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"
        WS3.Range("A1").AutoFilter Field:=6, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"
        WS4.Range("A1").AutoFilter Field:=6, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"
    
    Else 'other studies
        WS1.Name = "Form Status Report"
        Set WS2 = Sheets.Add
        WS2.Name = "Form Status Overview"
        WS2.Range("A1").Value = "Form Status"
        Call FormStatusOverview(WS1, WS2, 1, lastRow)
        WS2.Activate
        'Autofit and add borders for the form status overview table
        ActiveSheet.Range("A1").Select
        Call FormatTable
        
        WS1.Range("A1").AutoFilter Field:=6, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"
    End If
    
    Worksheets(1).Activate
    ActiveSheet.Range("A1").Select
    lastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row

    For Each xcell In ActiveSheet.Range("A1:F" & lastRow).Cells
        If xcell.Text = "Completed" Or xcell.Text = "Ready for Submission" Or xcell.Text = "Submitted to Sponsor" Then
            If SelectCells Is Nothing Then
                Set SelectCells = Range(xcell.Address)
            Else
                Set SelectCells = Union(SelectCells, Range(xcell.Address))
            End If
        End If
    Next
    If Not SelectCells Is Nothing Then
        SelectCells.EntireRow.Delete
    End If
    Set SelectCells = Nothing
    
    'sheet count
    SheetNum = Application.Sheets.Count
    For CurrentSheetNum = 2 To SheetNum
        Worksheets(CurrentSheetNum).Activate
        'remove filter
        Call RemoveFilter
        ActiveSheet.Range("A1").Select
        lastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row

        For Each xcell In ActiveSheet.Range("F1:F" & lastRow).Cells
            If xcell.Text = "Completed" Or xcell.Text = "Ready for Submission" Or xcell.Text = "Submitted to Sponsor" Then
                If SelectCells Is Nothing Then
                    Set SelectCells = Range(xcell.Address)
                Else
                    Set SelectCells = Union(SelectCells, Range(xcell.Address))
                End If
            End If
        Next
        
        If Not SelectCells Is Nothing Then
            SelectCells.EntireRow.Delete
        End If
        Set SelectCells = Nothing
        Call FormatTable
    Next CurrentSheetNum
    
    Sheets(2).Activate
    
    'Enable Screen Update
    Application.ScreenUpdating = True
    
    'Save file
    
    Dim modifiedDate As String
    modifiedDate = Now2Date(Now)
    Dim modifiedTime As String
    modifiedTime = Now2Time(Now)
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:=modifiedDate & "-XXXXX Form Status Report " & modifiedTime & " EST_Site.xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    If fileSaveName = False Then
        MsgBox "You haven't saved the document", vbExclamation
    Else
        ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
    End If
End Sub

Sub FormStatusOverview(WSNum1, WSNum2, i, lastRow)
    
    WSNum2.Range("B" & i).Value = "Count"
    WSNum2.Range("A" & i + 1).Value = "Completed"
    WSNum2.Range("A" & i + 2).Value = "Ready for Submission"
    WSNum2.Range("A" & i + 3).Value = "Submitted to Sponsor"
    WSNum2.Range("A" & i + 4).Value = "Work In Progress"
    WSNum2.Range("A" & i + 5).Value = "Incomplete"
    WSNum2.Range("B" & i + 1).Value = Application.WorksheetFunction.CountIf(WSNum1.Range("F2:K" & lastRow), "Completed")
    WSNum2.Range("B" & i + 2).Value = Application.WorksheetFunction.CountIf(WSNum1.Range("F2:K" & lastRow), "Ready for Submission")
    WSNum2.Range("B" & i + 3).Value = Application.WorksheetFunction.CountIf(WSNum1.Range("F2:K" & lastRow), "Submitted to Sponsor")
    WSNum2.Range("B" & i + 4).Value = Application.WorksheetFunction.CountIf(WSNum1.Range("F2:K" & lastRow), "Work In Progress")
    WSNum2.Range("B" & i + 5).Value = Application.WorksheetFunction.CountIf(WSNum1.Range("F2:K" & lastRow), "Incomplete")
End Sub
Sub QueryReportOverview(QueryWS, OverviewWS, TableStartRow)
    
    Dim lastRow As Long
    
    QueryWS.Activate
    QueryWS.Range("A1").Select
    lastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
    OverviewWS.Range("B" & TableStartRow).Value = "Count"
    'unique value
    OverviewWS.Range("A" & TableStartRow + 1).Value = "Resolved"
    OverviewWS.Range("A" & TableStartRow + 2).Value = "Closed"
    OverviewWS.Range("A" & TableStartRow + 3).Value = "Open"
    OverviewWS.Range("A" & TableStartRow + 4).Value = "Re-opened"
    'count
    OverviewWS.Range("B" & TableStartRow + 1).Value = Application.WorksheetFunction.CountIf(QueryWS.Range("L2:K" & lastRow), "Resolved")
    OverviewWS.Range("B" & TableStartRow + 2).Value = Application.WorksheetFunction.CountIf(QueryWS.Range("L2:K" & lastRow), "Closed")
    OverviewWS.Range("B" & TableStartRow + 3).Value = Application.WorksheetFunction.CountIf(QueryWS.Range("L2:K" & lastRow), "Open")
    OverviewWS.Range("B" & TableStartRow + 4).Value = Application.WorksheetFunction.CountIf(QueryWS.Range("L2:K" & lastRow), "Re-opened")
End Sub

Sub QuickReportQueryStatusFormat()
'
'QuickReportQueryStatusMacro Macro to reformat quick report patient form query status report, split into cohorts depend on studies
'
'Keyboard Shortcut: Ctrl+Shift+Q
'
Dim WS1 As Worksheet
Dim WS2 As Worksheet
Dim WS3 As Worksheet
Dim WS4 As Worksheet
Dim lastRow As Long
Dim WSCount As Integer
Dim Sh As Integer
Dim VSheet As Worksheet
Dim TempSheet As Worksheet
Dim xcell As Object
Dim SelectCells As Range
Dim fileSaveName As Variant


Application.ScreenUpdating = False

Set WS1 = Sheets(1)
WS1.Activate
WS1.Name = "All Cohorts Query Report"

'Counting LastRow and LastCol for studies which only have 1 output, Errl will skip over the code above and continue loop
WS1.Range("A1").Select
lastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
'Delete last row
Range("A" & lastRow).Select
Selection.EntireRow.Delete
lastRow = lastRow - 1
    
'Delete QUERY_STATUS_ID,PATIENT_ID,EVENT_NAME
Range("B:B,E:E,I:I").Select
Range("I1").Activate
Selection.Delete Shift:=xlToLeft



'Customization for each study
If WS1.Range("C2").Value = "827644" Then
    Call S14217.QQSR(WS1, WS2, WS3, WS4) '14217 study

ElseIf WS1.Range("C2").Value = "826085" Then
    Call S02916.QQSR(WS1, WS2, WS3, WS4) '02916 study
        
ElseIf WS1.Range("C2").Value = "850925" Then
    Call S01422.QQSR(WS1) '01422 Study
    
ElseIf WS1.Range("C2").Value = "823312" Then
    Call S15CT055.QQSR(WS1, WS2, WS3, WS4) '15CT055 study
    
ElseIf WS1.Range("C2").Value = "826250" Then
    Call S32816.QQSR(WS1, WS2, WS3, WS4) '32816 study

Else 'other studies
    WS1.Name = "Query Report"
    Set WS2 = Sheets.Add(Before:=WS1)
    WS2.Name = "Query Report Overview"
    WS2.Range("A1").Value = "Query Status"
    Call QueryReportOverview(WS1, WS2, 1)
    WS2.Activate
    'Autofit and add borders for the form status overview table
    WS2.Range("A1").Select
    Call FormatTable
End If

'Format table for all worksheets except the first one
WSCount = ActiveWorkbook.Worksheets.Count
For Sh = 2 To WSCount
    Sheets(Sh).Activate
    RemoveFilter
    Range("A1").Select
    FormatTable
    Sheets(Sh).Range("A1").AutoFilter Field:=12, Criteria1:="Open", Operator:=xlOr, Criteria2:="Re-opened"
Next Sh

'add validation tab
Sheets.Add After:=Sheets(Sheets.Count)
Set VSheet = Sheets(Sheets.Count)
VSheet.Name = "Validation tab"

'Filter to what is needed
WS1.Activate
RemoveFilter
With WS1.Range("A1")
    .AutoFilter Field:=12, Criteria1:="Open", Operator:=xlOr, Criteria2:="Re-opened"
    .AutoFilter Field:=8, Criteria1:=Array("Completed", "Ready for Submission", "Submitted to Sponsor"), Operator:=xlFilterValues
End With
WS1.Range("A1", WS1.Range("S1").End(xlDown)).SpecialCells(xlCellTypeVisible).Copy
VSheet.Range("D1").PasteSpecial
WS1.Range("A1").AutoFilter Field:=8


VSheet.Range("O1", VSheet.Range("O1").End(xlDown)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=VSheet.Range("A1"), Unique:=True
Dim VRow As Long
VRow = VSheet.Range("O1", VSheet.Range("O1").End(xlDown)).Rows.Count

VSheet.Activate
VSheet.Range("A1").Value = "Opened/Reopened with Completed/Ready for Submission/Submitted to Sponsor Status"

'Counting the number of instances the form is opened with Complete/Ready for submission status
Dim lngCount As Long
lngCount = Application.WorksheetFunction.CountA(Columns(1))
Dim cel As Range
If Not IsEmpty(VSheet.Range("A2")) Then
    For Each cel In VSheet.Range("A2:A" & lngCount).Cells
        VSheet.Range("B" & cel.Row).Value = Application.WorksheetFunction.CountIf(Range("O1:O" & VRow), cel.Text)
    Next cel
End If

'Format Validation Sheet
VSheet.Range("B1").Value = "Count of occurrence"
VSheet.Range("A1").Select
Call FormatTable

VSheet.Range("D1").Select
Call FormatTable
    
Sheets(2).Activate

Application.ScreenUpdating = True

'Save file
Dim modifiedDate As String
modifiedDate = Now2Date(Now)
Dim modifiedTime As String
modifiedTime = Now2Time(Now)

fileSaveName = Application.GetSaveAsFilename(InitialFileName:=modifiedDate & "-XXXXX Query Report " & modifiedTime & " EST.xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")

If fileSaveName = False Then
    MsgBox "You haven't saved the document", vbExclamation
Else
    ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
End If

End Sub

Sub VelosScheduleReview()

Dim WS1 As Worksheet
Dim WS2 As Worksheet
Dim lastRow As Long
Dim RowNum As Integer
Dim WS2Row As Integer
Dim arrSplitStrings As Variant
Dim singleString As Variant

Application.ScreenUpdating = False

Set WS1 = Sheets(1)
WS1.Activate

Set WS2 = Sheets.Add
WS2.Name = "Schedule Reviewing Tab"
WS2.Range("A1").Value = "Study Calendar Name"
WS2.Range("B1").Value = "Visit Name"
WS2.Range("C1").Value = "CRF Name"

WS1.Activate
WS1.Range("A1").Select
lastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row

WS2Row = 2
For RowNum = 9 To lastRow
    
    If Left(WS1.Range("B" & RowNum).Value, 6) = "Visit:" Then
        arrSplitStrings = Split(WS1.Range("E" & RowNum + 1).Value, ", ")
        For Each singleString In arrSplitStrings
            
            WS1.Range("B2").Copy WS2.Range("A" & WS2Row)
            WS2.Range("B" & WS2Row).Value = Trim(Mid(WS1.Range("B" & RowNum).Value, 8))
            WS2.Range("C" & WS2Row).Value = Trim(singleString)
            WS2Row = WS2Row + 1
        Next singleString
        
    
    End If
    
Next RowNum
WS2.Activate
'remove ASCII 160
ActiveSheet.Cells.Replace Chr(160), ""

FormatTable
WS2.Range("A1").AutoFilter

Application.ScreenUpdating = True

End Sub

