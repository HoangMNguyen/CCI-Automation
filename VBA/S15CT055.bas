Attribute VB_Name = "S15CT055"
Sub PRIOR_ANP() '15CT055- Prior Antineoplastic (ANP) Therapy V3

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
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Filtered Prior ANP"
    Set WS2 = Sheets("Filtered Prior ANP")
End With
'Set up header for WS2
WS1.Range("C8").Copy WS2.Range("A1")
For i = 16 To 23 'was 25, need to fix the production site
    WS1.Cells(8, i).Copy WS2.Cells(1, i - 14)
Next i

'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
'Loop rows
l = 1
For i = 9 To RowNum
    For j = 16 To 175 Step 8
    If WS1.Cells(i, j).Value <> "No" Then
        l = l + 1
        WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
        For k = 0 To 7
                WS1.Cells(i, k + j).Copy WS2.Cells(l, k + 2)
        Next k
    End If
    Next j
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat
End Sub

Sub QFSR(WS1, WS2, WS3, WS4, LastRow)

    'Copy cohort 2 data to another tab
     WS1.Range("A1").AutoFilter Field:=3, Criteria1:=Array( _
        "15CT055 Cohort 2 LTFU Calendar", _
        "15CT055 Cohort 2 Primary Calendar", _
        "15CT055 Cohort 2 Retreatment Calendar V2", _
        "15CT055 Cohort 2 Retx-LTFU Calendar", _
        "15CT055 Exception Retreatment Calendar"), Operator:=xlFilterValues
   
    WS1.Range("A1:J" & LastRow).Select
    Selection.Copy
    
    Set WS2 = Sheets.Add
    WS2.Name = "Cohort 2 Form Status"
    WS2.Paste
    Call Main.OutFormat

   Sheets("All Cohorts Form Status Report").Move Before:=Sheets(1)
   
    'Copy cohort 1 data to another tab
    Sheets("All Cohorts Form Status Report").Select
    WS1.Range("A1").AutoFilter Field:=3, Criteria1:="15CT055 Calendar", Operator:=xlOr, Criteria2:="15CT055-LTFU Calendar"
    WS1.Range("A1:J" & LastRow).Select
    Selection.Copy
    
    Set WS3 = Sheets.Add
    WS3.Name = "Cohort 1 Form Status"
    WS3.Paste
    Call Main.OutFormat
 
    Sheets("All Cohorts Form Status Report").Move Before:=Sheets(1)

    Set WS4 = Sheets.Add
    WS4.Name = "Form Status Overview"
    
    WS4.Range("A1").Value = "All Cohorts Form Status"
    Call FormStatusOverview(WS1, WS4, 1, LastRow)
    
    WS4.Range("A8").Value = "Cohort 1 Form Status"
    Call FormStatusOverview(WS3, WS4, 8, LastRow)
    
    WS4.Range("A15").Value = "Cohort 2 Form Status"
    Call FormStatusOverview(WS2, WS4, 15, LastRow)
 
    'Autofit and add borders for the form status overview table
    ActiveSheet.Range("A1").Select
    Call FormatTable

    ActiveSheet.Range("A8").Select
    Call FormatTable

    ActiveSheet.Range("A15").Select
    Call FormatTable

    WS1.Range("A1").AutoFilter Field:=3

End Sub

Sub QQSR(WS1, WS2, WS3, WS4)

Dim LastRow As Long

Set WS2 = Sheets.Add(After:=WS1)
WS2.Name = "Cohort 2 Query Report"

Set WS3 = Sheets.Add(After:=WS2)
WS3.Name = "Cohort 1 Query Report"

Set WS4 = Sheets.Add(Before:=WS1)
WS4.Name = "Query Report Overview"

WS1.Activate
WS1.Range("A1").Select
LastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row

'Copy cohort 2 data to another tab
WS1.Range("A1").AutoFilter Field:=5, Criteria1:=Array( _
   "15CT055 Cohort 2 LTFU Calendar", _
   "15CT055 Cohort 2 Primary Calendar", _
   "15CT055 Cohort 2 Retreatment Calendar V2", _
   "15CT055 Cohort 2 Retx-LTFU Calendar", _
   "15CT055 Exception Retreatment Calendar"), Operator:=xlFilterValues

WS1.Range("A1:S" & LastRow).SpecialCells(xlCellTypeVisible).Copy
WS2.Paste

'Copy cohort 1 data to another tab
WS1.Range("A1").AutoFilter Field:=5, Criteria1:="15CT055 Calendar", Operator:=xlOr, Criteria2:="15CT055-LTFU Calendar"
WS1.Range("A1:S" & LastRow).SpecialCells(xlCellTypeVisible).Copy
WS3.Paste

WS4.Range("A1").Value = "All Cohorts Query Status"
Call QueryReportOverview(WS1, WS4, 1)

WS4.Range("A7").Value = "Cohort 1 Query Status"
Call QueryReportOverview(WS3, WS4, 7)

WS4.Range("A13").Value = "Cohort 2 Query Status"
Call QueryReportOverview(WS2, WS4, 13)

'Autofit and add borders for the form status overview table
WS4.Activate
WS4.Range("A1").Select
FormatTable

WS4.Range("A7").Select
FormatTable

WS4.Range("A13").Select
FormatTable

End Sub
