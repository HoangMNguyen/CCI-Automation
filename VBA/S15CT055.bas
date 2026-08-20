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
    .Sheets.Add(After:=.Sheets(.Sheets.count)).Name = "Filtered Prior ANP"
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
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).row
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

Sub QFSR(WS1, WS2, WS3, WS4, WS5, WS6, lastRow)

    Dim splitRow As Long, splitLast As Long
    Dim splitSubj As String
    Dim dropRows As Range

    'Copy cohort 3A data to another tab, cohort 3A shares the cohort 3 calendars
     WS1.Range("A1").AutoFilter Field:=3, Criteria1:=Array( _
        "15CT055 Cohort 3 LTFU Calendar", _
        "15CT055 Cohort 3 Primary Calendar", _
        "15CT055 Cohort 3 Retreatment Calendar", _
        "15CT055 Cohort 3 Retx-LTFU Calendar"), Operator:=xlFilterValues

    WS1.Range("A1:J" & lastRow).Select
    Selection.Copy

    Set WS6 = Sheets.Add
    WS6.Name = "Cohort 3A Form Status"
    WS6.Paste
    Call Main.OutFormat

    'Drop subjects below 47, they belong on the cohort 3 tab
    splitLast = WS6.Cells(WS6.Rows.count, 2).End(xlUp).row
    For splitRow = 2 To splitLast
        splitSubj = CStr(WS6.Cells(splitRow, 2).Value)
        If Val(Mid$(splitSubj, InStrRev(splitSubj, "-") + 1)) < 47 Then
            If dropRows Is Nothing Then
                Set dropRows = WS6.Rows(splitRow)
            Else
                Set dropRows = Union(dropRows, WS6.Rows(splitRow))
            End If
        End If
    Next splitRow
    If Not dropRows Is Nothing Then dropRows.Delete
    Set dropRows = Nothing

   Sheets("All Cohorts Form Status Report").Move Before:=Sheets(1)

    'Copy cohort 3 data to another tab
     WS1.Range("A1").AutoFilter Field:=3, Criteria1:=Array( _
        "15CT055 Cohort 3 LTFU Calendar", _
        "15CT055 Cohort 3 Primary Calendar", _
        "15CT055 Cohort 3 Retreatment Calendar", _
        "15CT055 Cohort 3 Retx-LTFU Calendar"), Operator:=xlFilterValues

    WS1.Range("A1:J" & lastRow).Select
    Selection.Copy

    Set WS5 = Sheets.Add
    WS5.Name = "Cohort 3 Form Status"
    WS5.Paste
    Call Main.OutFormat

    'Drop subject 47 and above, they belong on the cohort 3A tab
    splitLast = WS5.Cells(WS5.Rows.count, 2).End(xlUp).row
    For splitRow = 2 To splitLast
        splitSubj = CStr(WS5.Cells(splitRow, 2).Value)
        If Val(Mid$(splitSubj, InStrRev(splitSubj, "-") + 1)) >= 47 Then
            If dropRows Is Nothing Then
                Set dropRows = WS5.Rows(splitRow)
            Else
                Set dropRows = Union(dropRows, WS5.Rows(splitRow))
            End If
        End If
    Next splitRow
    If Not dropRows Is Nothing Then dropRows.Delete
    Set dropRows = Nothing

   Sheets("All Cohorts Form Status Report").Move Before:=Sheets(1)

    'Copy cohort 2 data to another tab
     WS1.Range("A1").AutoFilter Field:=3, Criteria1:=Array( _
        "15CT055 Cohort 2 LTFU Calendar", _
        "15CT055 Cohort 2 Primary Calendar", _
        "15CT055 Cohort 2 Retreatment Calendar V2", _
        "15CT055 Cohort 2 Retx-LTFU Calendar", _
        "15CT055 Exception Retreatment Calendar"), Operator:=xlFilterValues
   
    WS1.Range("A1:J" & lastRow).Select
    Selection.Copy
    
    Set WS2 = Sheets.Add
    WS2.Name = "Cohort 2 Form Status"
    WS2.Paste
    Call Main.OutFormat

   Sheets("All Cohorts Form Status Report").Move Before:=Sheets(1)
   
    'Copy cohort 1 data to another tab
    Sheets("All Cohorts Form Status Report").Select
    WS1.Range("A1").AutoFilter Field:=3, Criteria1:="15CT055 Calendar", Operator:=xlOr, Criteria2:="15CT055-LTFU Calendar"
    WS1.Range("A1:J" & lastRow).Select
    Selection.Copy
    
    Set WS3 = Sheets.Add
    WS3.Name = "Cohort 1 Form Status"
    WS3.Paste
    Call Main.OutFormat
 
    Sheets("All Cohorts Form Status Report").Move Before:=Sheets(1)

    Set WS4 = Sheets.Add
    WS4.Name = "Form Status Overview"
    
    WS4.Range("A1").Value = "All Cohorts Form Status"
    Call FormStatusOverview(WS1, WS4, 1, lastRow)
    
    WS4.Range("A8").Value = "Cohort 1 Form Status"
    Call FormStatusOverview(WS3, WS4, 8, lastRow)
    
    WS4.Range("A15").Value = "Cohort 2 Form Status"
    Call FormStatusOverview(WS2, WS4, 15, lastRow)
    
    WS4.Range("A22").Value = "Cohort 3 Form Status"
    Call FormStatusOverview(WS5, WS4, 22, lastRow)

    WS4.Range("A29").Value = "Cohort 3A Form Status"
    Call FormStatusOverview(WS6, WS4, 29, lastRow)

    'Autofit and add borders for the form status overview table
    ActiveSheet.Range("A1").Select
    Call FormatTable

    ActiveSheet.Range("A8").Select
    Call FormatTable

    ActiveSheet.Range("A15").Select
    Call FormatTable
    
    ActiveSheet.Range("A22").Select
    Call FormatTable

    ActiveSheet.Range("A29").Select
    Call FormatTable

    WS1.Range("A1").AutoFilter Field:=3

End Sub

Sub QQSR(WS1 As Worksheet, WS2 As Worksheet, WS3 As Worksheet, WS4 As Worksheet, WS5 As Worksheet, WS6 As Worksheet)

Dim lastRow As Long
Dim splitRow As Long, splitLast As Long
Dim splitSubj As String
Dim dropRows As Range

Set WS5 = Sheets.Add(After:=WS1)
WS5.Name = "Cohort 3 Query Report"

Set WS6 = Sheets.Add(After:=WS5)
WS6.Name = "Cohort 3A Query Report"

Set WS2 = Sheets.Add(After:=WS6)
WS2.Name = "Cohort 2 Query Report"

Set WS3 = Sheets.Add(After:=WS2)
WS3.Name = "Cohort 1 Query Report"

Set WS4 = Sheets.Add(Before:=WS1)
WS4.Name = "Query Report Overview"

WS1.Activate
WS1.Range("A1").Select
lastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious).row
Dim calendarNameCol As Integer
calendarNameCol = L2N(FindColumn(WS1, "CALENDAR_NAME"))
'Copy cohort 2 data to another tab
WS1.Range("A1").AutoFilter Field:=calendarNameCol, Criteria1:=Array( _
   "15CT055 Cohort 2 LTFU Calendar", _
   "15CT055 Cohort 2 Primary Calendar", _
   "15CT055 Cohort 2 Retreatment Calendar V2", _
   "15CT055 Cohort 2 Retx-LTFU Calendar", _
   "15CT055 Exception Retreatment Calendar"), Operator:=xlFilterValues

WS1.Range("A1:T" & lastRow).SpecialCells(xlCellTypeVisible).Copy
WS2.Paste

'Copy cohort 1 data to another tab
WS1.Range("A1").AutoFilter Field:=calendarNameCol, Criteria1:="15CT055 Calendar", Operator:=xlOr, Criteria2:="15CT055-LTFU Calendar"
WS1.Range("A1:T" & lastRow).SpecialCells(xlCellTypeVisible).Copy
WS3.Paste

'Copy cohort 3 and cohort 3A data to another tab, cohort 3A shares the cohort 3 calendars
WS1.Range("A1").AutoFilter Field:=calendarNameCol, Criteria1:=Array( _
        "15CT055 Cohort 3 LTFU Calendar", _
        "15CT055 Cohort 3 Primary Calendar", _
        "15CT055 Cohort 3 Retreatment Calendar", _
        "15CT055 Cohort 3 Retx-LTFU Calendar"), Operator:=xlFilterValues

WS1.Range("A1:T" & lastRow).SpecialCells(xlCellTypeVisible).Copy
WS5.Paste

'Cohort 3A shares the cohort 3 calendars, duplicate the rows then split them by subject number
WS5.UsedRange.Copy Destination:=WS6.Range("A1")

'Drop subject 47 and above from cohort 3, they belong on the cohort 3A tab
splitLast = WS5.Cells(WS5.Rows.count, 4).End(xlUp).row
For splitRow = 2 To splitLast
    splitSubj = CStr(WS5.Cells(splitRow, 4).Value)
    If Val(Mid$(splitSubj, InStrRev(splitSubj, "-") + 1)) >= 47 Then
        If dropRows Is Nothing Then
            Set dropRows = WS5.Rows(splitRow)
        Else
            Set dropRows = Union(dropRows, WS5.Rows(splitRow))
        End If
    End If
Next splitRow
If Not dropRows Is Nothing Then dropRows.Delete
Set dropRows = Nothing

'Drop subjects below 47 from cohort 3A, they belong on the cohort 3 tab
splitLast = WS6.Cells(WS6.Rows.count, 4).End(xlUp).row
For splitRow = 2 To splitLast
    splitSubj = CStr(WS6.Cells(splitRow, 4).Value)
    If Val(Mid$(splitSubj, InStrRev(splitSubj, "-") + 1)) < 47 Then
        If dropRows Is Nothing Then
            Set dropRows = WS6.Rows(splitRow)
        Else
            Set dropRows = Union(dropRows, WS6.Rows(splitRow))
        End If
    End If
Next splitRow
If Not dropRows Is Nothing Then dropRows.Delete
Set dropRows = Nothing

WS1.Activate
RemoveFilter

WS4.Range("A1").Value = "All Cohorts Query Status"
Call QueryReportOverview(WS1, WS4, 1)

WS4.Range("A7").Value = "Cohort 1 Query Status"
Call QueryReportOverview(WS3, WS4, 7)

WS4.Range("A13").Value = "Cohort 2 Query Status"
Call QueryReportOverview(WS2, WS4, 13)

WS4.Range("A19").Value = "Cohort 3 Query Status"
Call QueryReportOverview(WS5, WS4, 19)

WS4.Range("A25").Value = "Cohort 3A Query Status"
Call QueryReportOverview(WS6, WS4, 25)

'Autofit and add borders for the form status overview table
WS4.Activate
WS4.Range("A1").Select
FormatTable

WS4.Range("A7").Select
FormatTable

WS4.Range("A13").Select
FormatTable

WS4.Range("A19").Select
FormatTable

WS4.Range("A25").Select
FormatTable

End Sub
