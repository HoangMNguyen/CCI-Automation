Attribute VB_Name = "S32816"
Sub PRIOR_ONC() 'Prior Oncology Therapy V1 (must run the report in linear format, otherwise not working)

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
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Filtered Prior ONC"
    Set WS2 = Sheets("Filtered Prior ONC")
End With
'Set up header for WS2
WS1.Range("C11").Copy WS2.Range("A1")
WS1.Range("P11").Copy WS2.Range("B1")
WS2.Range("C1").Value = "Regimen number"
WS2.Range("D1").Value = "Medication"
WS2.Range("E1").Value = "Therapy Type"
WS2.Range("F1").Value = "Other, specify"
WS2.Range("G1").Value = "Start date of first dose"
WS2.Range("H1").Value = "End date of last dose"
WS2.Range("I1").Value = "Number of cycles"

'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
'Loop rows
l = 1
For i = 12 To RowNum
    For j = 0 To 5
        l = l + 1
        WS1.Range("C" & i).Copy WS2.Range("A" & (l))
        WS1.Range("P" & i).Copy WS2.Range("B" & (l))
        For k = 17 To 58 Step 6
            WS1.Cells(i, j + k).Copy WS2.Cells(l, ((k - 10) / 6 + 2))
            
        Next k
    Next j
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat

End Sub
Sub QFSR(WS1, WS2, WS3, WS4, LastRow)

    'Copy cohort -3 data to another tab
    WS1.Range("A1").AutoFilter Field:=3, Criteria1:=Array( _
        "32816 Cohort -3 Primary Calendar", _
        "32816 Cohort -3 LTFU Calendar", _
        "32816-08 Primary Retreatment Calendar"), Operator:=xlFilterValues

    WS1.Range("A1:J" & LastRow).Select
    Selection.Copy
    
    Set WS2 = Sheets.Add
    WS2.Name = "Cohort -3 Form Status"
    WS2.Paste
    Call Main.OutFormat

   Sheets("All Cohorts Form Status Report").Move Before:=Sheets(1)
   
    'Copy cohorts 1-3 data to another tab
    Sheets("All Cohorts Form Status Report").Select
    WS1.Range("A1").AutoFilter Field:=3, Criteria1:="32816 Calendar", Operator:=xlOr, Criteria2:="32816 LTFU Calendar"
    WS1.Range("A1:J" & LastRow).Select
    Selection.Copy
    
    Set WS3 = Sheets.Add
    WS3.Name = "Cohort 1-3 Form Status"
    WS3.Paste
    Call Main.OutFormat
    
    'Copy cohorts 4 data to another tab
    Sheets("All Cohorts Form Status Report").Select
    WS1.Range("A1").AutoFilter Field:=3, Criteria1:="32816 Cohort 4 Primary Calendar", Operator:=xlOr, Criteria2:="32816 Cohort 4 LTFU Calendar"
    WS1.Range("A1:J" & LastRow).Select
    Selection.Copy
    
    Set WS4 = Sheets.Add(After:=WS2)
    WS4.Name = "Cohort 4 Form Status"
    WS4.Paste
    Call Main.OutFormat
 
    Sheets("All Cohorts Form Status Report").Move Before:=Sheets(1)

    Set WS5 = Sheets.Add
    WS5.Name = "Form Status Overview"
    
    WS5.Range("A1").Value = "All Cohorts Form Status"
    Call FormStatusOverview(WS1, WS5, 1, LastRow)
    
    WS5.Range("A8").Value = "Cohorts 1-3 Form Status"
    Call FormStatusOverview(WS3, WS5, 8, LastRow)
    
    WS5.Range("A15").Value = "Cohort -3 Form Status"
    Call FormStatusOverview(WS2, WS5, 15, LastRow)
    
    WS5.Range("A22").Value = "Cohort 4 Form Status"
    Call FormStatusOverview(WS4, WS5, 22, LastRow)
 
    'Autofit and add borders for the form status overview table
    ActiveSheet.Range("A1").Select
    Call FormatTable

    ActiveSheet.Range("A8").Select
    Call FormatTable

    ActiveSheet.Range("A15").Select
    Call FormatTable
    
    ActiveSheet.Range("A22").Select
    Call FormatTable

    WS1.Range("A1").AutoFilter Field:=3


End Sub

Sub QQSR(WS1, WS2, WS3, WS4)

Dim LastRow As Long

Set WS2 = Sheets.Add(After:=WS1)
WS2.Name = "Cohort 1-3 Query Report"

Set WS3 = Sheets.Add(After:=WS2)
WS3.Name = "Cohort -3 Query Report"

Set WS4 = Sheets.Add(After:=WS3)
WS4.Name = "Cohort 4 Query Report"

Set WS5 = Sheets.Add(Before:=WS1)
WS5.Name = "Query Report Overview"

WS1.Activate
WS1.Range("A1").Select
LastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row

'Copy cohorts 1-3 data to another tab
WS1.Range("A1").AutoFilter Field:=5, Criteria1:="32816 Calendar", Operator:=xlOr, Criteria2:="32816 LTFU Calendar"
WS1.Range("A1:S" & LastRow).SpecialCells(xlCellTypeVisible).Copy
WS2.Paste


'Copy cohort 2&3 data to another tab
WS1.Range("A1").AutoFilter Field:=5, Criteria1:=Array( _
"32816 Cohort -3 Primary Calendar", _
"32816 Cohort -3 LTFU Calendar", _
"32816-08 Primary Retreatment Calendar"), Operator:=xlFilterValues
WS1.Range("A1:S" & LastRow).SpecialCells(xlCellTypeVisible).Copy
WS3.Paste


'Copy cohorts 4 data to another tab
WS1.Range("A1").AutoFilter Field:=5, Criteria1:="32816 Cohort 4 Primary Calendar", Operator:=xlOr, Criteria2:="32816 Cohort 4 LTFU Calendar"
WS1.Range("A1:S" & LastRow).SpecialCells(xlCellTypeVisible).Copy
WS4.Paste


WS1.Activate
RemoveFilter

WS5.Range("A1").Value = "All Cohorts Query Status"
Call QueryReportOverview(WS1, WS5, 1)

WS5.Range("A7").Value = "Cohort 1-3 Query Status"
Call QueryReportOverview(WS2, WS5, 7)

WS5.Range("A13").Value = "Cohort -3 Query Status"
Call QueryReportOverview(WS3, WS5, 13)

WS5.Range("A19").Value = "Cohort 4 Query Status"
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


