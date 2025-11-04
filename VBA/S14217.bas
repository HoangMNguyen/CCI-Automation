Attribute VB_Name = "S14217"
Sub PRIOR_ANP() 'reformat Prior Antineoplastic (ANP) Therapy V2

Dim WS1 As Worksheet
Dim WS2 As Worksheet
Dim WB1 As Workbook
Dim RowNum As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer

Set WB1 = ActiveWorkbook
Set WS1 = ActiveSheet
'Add WS2 as output
With WB1
    .Sheets.Add(After:=.Sheets(.Sheets.count)).Name = "Filtered Prior ANP"
    Set WS2 = Sheets("Filtered Prior ANP")
End With
'Set up header for WS2
WS1.Range("C8").Copy WS2.Range("A1")
For i = 16 To 25
    WS1.Cells(8, i).Copy WS2.Cells(1, i - 14)
Next i



'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).row
'Loop rows
l = 1
For i = 9 To RowNum
    For j = 16 To 115 Step 10
    If WS1.Cells(i, j).Value <> "No" Then
        l = l + 1
        WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
        For k = 0 To 9
                WS1.Cells(i, k + j).Copy WS2.Cells(l, k + 2)
        Next k
    End If
    Next j
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat

End Sub

Sub MEDHX() 'reformat Medical History V2

Dim WS1 As Worksheet
Dim WS2 As Worksheet
Dim WB1 As Workbook
Dim RowNum As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer

Set WB1 = ActiveWorkbook
Set WS1 = ActiveSheet
'Add WS2 as output
With WB1
    .Sheets.Add(After:=.Sheets(.Sheets.count)).Name = "Filtered Medical History"
    Set WS2 = Sheets("Filtered Medical History")
End With
'Set up header for WS2
WS1.Range("C8").Copy WS2.Range("A1")
For i = 16 To 25
    WS1.Cells(8, i).Copy WS2.Cells(1, i - 14)
Next i



'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).row
'Loop rows
l = 1
For i = 9 To RowNum
    For j = 16 To 115 Step 10
    If WS1.Cells(i, j).Value <> "No" Then
        l = l + 1
        WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
        For k = 0 To 9
                WS1.Cells(i, k + j).Copy WS2.Cells(l, k + 2)
        Next k
    End If
    Next j
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat

End Sub

Sub QFSR(WS1, WS2, WS3, WS4, lastRow)

    'Copy cohort 1 data to another tab
    Sheets("All Cohorts Form Status Report").Select
    Set WS1 = Sheets("All Cohorts Form Status Report")
    WS1.Range("A1").AutoFilter Field:=3, Criteria1:="14217 Subject Study Calendar", Operator:=xlOr, Criteria2:="14217 Subject Study Calendar V2"
    WS1.Range("A1:J" & lastRow).Select
    Selection.Copy
    
    Set WS2 = Sheets.Add(After:=Sheets(Sheets.count))
    WS2.Name = "Cohort 1 Form Status"
    WS2.Paste
    Call Main.OutFormat
    
    Sheets("All Cohorts Form Status Report").Move Before:=Sheets(1)

    'Copy cohort 2&3 data to another tab
    WS1.Range("A1").AutoFilter Field:=3, Criteria1:="14217 Subject Study Calendar - Local Delivery", Operator:=xlOr, Criteria2:="14217 Subject Study Calendar V2 - Local Delivery"
    WS1.Range("A1:J" & lastRow).Select
    Selection.Copy
    
    Set WS3 = Sheets.Add(After:=Sheets(Sheets.count))
    WS3.Name = "Cohorts 2&3 Form Status"
    WS3.Paste
    Call Main.OutFormat
    
    Sheets("All Cohorts Form Status Report").Move Before:=Sheets(1)

    'Copy cohort 4 data to another tab
    WS1.Range("A1").AutoFilter Field:=3, Criteria1:="14217 Subject Study Calendar V3 - Local Delivery"
    WS1.Range("A1:J" & lastRow).Select
    Selection.Copy
    
    Set WS4 = Sheets.Add(After:=Sheets(Sheets.count))
    WS4.Name = "Cohorts 4 Form Status"
    WS4.Paste
    Call Main.OutFormat

    Sheets("All Cohorts Form Status Report").Move Before:=Sheets(1)
    
    Set WS5 = Sheets.Add
    WS5.Name = "Form Status Overview"
    
    WS5.Range("A1").Value = "All Cohorts Form Status"
    Call FormStatusOverview(WS1, WS5, 1, lastRow)
    
    WS5.Range("A8").Value = "Cohort 1 Form Status"
    Call FormStatusOverview(WS2, WS5, 8, lastRow)
   
    
    WS5.Range("A15").Value = "Cohorts 2&3 Form Status"
    Call FormStatusOverview(WS3, WS5, 15, lastRow)
    
    WS5.Range("A22").Value = "Cohorts 4 Form Status"
    Call FormStatusOverview(WS4, WS5, 22, lastRow)

    'Autofit and add borders for the form status overview table
    ActiveSheet.Range("A1").Select
    Call FormatTable
    
    ActiveSheet.Range("A8").Select
    Call FormatTable

    ActiveSheet.Range("A15").Select
    Call FormatTable
    
    ActiveSheet.Range("A22").Select
    Call FormatTable
    
    WS1.Range("A1").AutoFilter Field:=3 'All cohorts sheet select all calendars

End Sub

Sub QQSR(WS1 As Worksheet, WS2 As Worksheet, WS3 As Worksheet, WS4 As Worksheet)

Dim lastRow As Long
Dim WS5 As Worksheet

Set WS3 = Sheets.Add(After:=WS1)
WS3.Name = "Cohort 1 Query Report"

Set WS2 = Sheets.Add(After:=WS3)
WS2.Name = "Cohorts 2&3 Query Report"

Set WS4 = Sheets.Add(After:=WS2)
WS4.Name = "Cohort 4 Query Report"

Set WS5 = Sheets.Add(Before:=WS1)
WS5.Name = "Query Report Overview"

WS1.Activate
WS1.Range("A1").Select
lastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious).row
Dim calendarNameCol As Integer
calendarNameCol = L2N(FindColumn(WS1, "CALENDAR_NAME"))
'Copy cohort 2&3 data to another tab
WS1.Range("A1").AutoFilter Field:=calendarNameCol, Criteria1:="14217 Subject Study Calendar - Local Delivery", Operator:=xlOr, Criteria2:="14217 Subject Study Calendar V2 - Local Delivery"
WS1.Range("A1:T" & lastRow).SpecialCells(xlCellTypeVisible).Copy
WS2.Paste

'Copy cohort 1 data to another tab
WS1.Range("A1").AutoFilter Field:=calendarNameCol, Criteria1:="14217 Subject Study Calendar", Operator:=xlOr, Criteria2:="14217 Subject Study Calendar V2"
WS1.Range("A1:T" & lastRow).SpecialCells(xlCellTypeVisible).Copy
WS3.Paste

'Copy cohort 4 data to another tab
WS1.Range("A1").AutoFilter Field:=calendarNameCol, Criteria1:="14217 Subject Study Calendar V3 - Local Delivery"
WS1.Range("A1:T" & lastRow).SpecialCells(xlCellTypeVisible).Copy
WS4.Paste

WS1.Activate
RemoveFilter

WS5.Range("A1").Value = "All Cohorts Query Status"
Call QueryReportOverview(WS1, WS5, 1)

WS5.Range("A7").Value = "Cohort 1 Query Status"
Call QueryReportOverview(WS3, WS5, 7)

WS5.Range("A13").Value = "Cohorts 2&3 Query Status"
Call QueryReportOverview(WS2, WS5, 13)

WS5.Range("A19").Value = "Cohorts 4 Query Status"
Call QueryReportOverview(WS4, WS5, 19)

'Autofit and add borders for the form status overview table
WS5.Activate
WS5.Range("A1").Select
FormatTable

WS5.Range("A7").Select
FormatTable

WS5.Range("A13").Select
FormatTable

WS5.Range("A19").Select
FormatTable

End Sub
