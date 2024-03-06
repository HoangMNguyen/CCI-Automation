Attribute VB_Name = "ConmedAE"
Option Explicit

Sub FormatConmed()


Dim WS1 As Worksheet
Dim WS2 As Worksheet
Dim RowNum As Integer
Dim FSField As Long
Dim FSFL As String
Dim UniqueID As Range
Dim i As Integer
Dim lastRow As Long
Dim RemoveRow As Long

Set WS1 = ActiveSheet
    
'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
lastRow = ActiveSheet.Cells.Find(What:="*", SearchDirection:=xlPrevious).Row

If WS1.Range("B5").Value = "Basic Stats" Then 'report without header
    'delete first 7 rows and last 2 rows.
    Rows("1:7").Select
    Selection.Delete Shift:=xlUp
    Range("A" & (lastRow - 6) & ":A" & (lastRow - 5)).Select 'last row was returned before deletion of first 7 rows
ElseIf WS1.Range("B8").Value = "Basic Stats" Then 'report with header
'delete first 10 rows and last 2 rows.
    Rows("1:10").Select
    Selection.Delete Shift:=xlUp
    Range("A" & (lastRow - 9) & ":A" & (lastRow - 8)).Select 'last row was returned before deletion of first 10 rows
End If
Selection.EntireRow.Delete
ActiveSheet.Name = "Conmed Status"
 
    
'Adds Conmed Overview Sheet
FSField = Application.Match("Form Status", WS1.Range("A1:AF1"), 0)
Set WS2 = Sheets.Add
    WS2.Name = "Conmed Overview"
    FSFL = Split(Cells(1, FSField).Address, "$")(1) 'split ($G$1","$")return the address of row 1 and column form status exisits (column G)
   ' MsgBox (Cells(1, FSField).Address)
    Sheets("Conmed Status").Columns(FSFL & ":" & FSFL).Copy
    Sheets("Conmed Overview").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

'Removes Duplicates from Form Conmed Overview Tab, formats and removes blanks
    Set UniqueID = WS2.Range("A1")
    WS1.Range(FSFL & ":" & FSFL).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=UniqueID, Unique:=True 'copy unique values of form status to conmed overview column A
    Range("B1").Value = "Totals" 'assign header value
    WS1.Range(FSFL & ":" & FSFL).Copy WS2.Range("C:C") 'copy everything from conmed status column G to conmed overview column C
   ' MsgBox (FSFL)
    For i = 2 To 10 'form status has less than 10 types, use 10 is enough
        WS2.Range("B" & i).Value = WorksheetFunction.CountIf(WS2.Range("C1:C" & lastRow), Range("A" & i).Value)
    Next i
    WS2.Range("C:C").Delete
    
    Range("A11:B" & lastRow).ClearFormats
    Columns("A:A").Select
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    Columns("A:A").ColumnWidth = 18.5
    Columns("B:B").ColumnWidth = 18.5
    
    Range("A1:B1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$B$10").AutoFilter Field:=1, Criteria1:="<>"
    
'Format the Conmed Status Report Sheet
    WS1.Activate
    
    ActiveSheet.Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.WrapText = True
    Selection.VerticalAlignment = xlCenter
    Selection.HorizontalAlignment = xlLeft
    Selection.Borders.LineStyle = xlContinuous
    Selection.Borders.Weight = xlThin
    Selection.AutoFilter
    ActiveSheet.Range("A1").AutoFilter Field:=FSField, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"
End Sub

Sub MainAE()

Dim WS1 As Worksheet
Dim WS2 As Worksheet
Dim RowNum As Integer
Dim FSField As Long
Dim FSFL As String
Dim UniqueID As Range
Dim i As Integer
Dim lastRow As Long
Dim RemoveRow As Long

'Disable Screen Update
Application.ScreenUpdating = False
Set WS1 = ActiveSheet
    
'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
lastRow = ActiveSheet.Cells.Find(What:="*", SearchDirection:=xlPrevious).Row



If WS1.Range("B5").Value = "Basic Stats" Then 'report without header
    'delete first 7 rows and last 2 rows.
    Rows("1:7").Select
    Selection.Delete Shift:=xlUp
    Range("A" & (lastRow - 6) & ":A" & (lastRow - 5)).Select 'last row was returned before deletion of first 7 rows
ElseIf WS1.Range("B8").Value = "Basic Stats" Then 'report with header
'delete first 10 rows and last 2 rows.
    Rows("1:10").Select
    Selection.Delete Shift:=xlUp
    Range("A" & (lastRow - 9) & ":A" & (lastRow - 8)).Select 'last row was returned before deletion of first 10 rows
End If
Selection.EntireRow.Delete
ActiveSheet.Name = "Adverse Event Status"
 
    
On Error GoTo Err1:
FSField = Application.Match("Form Status", WS1.Range("A1:CQ1"), 0) 'for CCI AE
'MsgBox (FSField)
Err1:
On Error Resume Next
FSField = Application.Match("AE Status", WS1.Range("A1:CQ1"), 0) 'for AE page


Set WS2 = Sheets.Add
WS2.Name = "Adverse Event Overview"
FSFL = Split(Cells(1, FSField).Address, "$")(1)
' MsgBox (Cells(1, FSField).Address)
Sheets("Adverse Event Status").Columns(FSFL & ":" & FSFL).Copy
Sheets("Adverse Event Overview").Activate
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False

'Removes Duplicates from Form Adverse Event Overview Tab, formats and removes blanks
 Set UniqueID = WS2.Range("A1")
 WS1.Range(FSFL & ":" & FSFL).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=UniqueID, Unique:=True 'copy unique values of form status to Adverse Event overview column A
 Range("B1").Value = "Totals" 'assign header value
 WS1.Range(FSFL & ":" & FSFL).Copy WS2.Range("C:C") 'copy everything from Adverse Event status column G to Adverse Event overview column C
' MsgBox (FSFL)
For i = 2 To 10 'form status has less than 10 types, use 10 is enough
    WS2.Range("B" & i).Value = WorksheetFunction.CountIf(WS2.Range("C1:C" & lastRow), Range("A" & i).Value)
Next i
WS2.Range("C:C").Delete

Range("A11:B" & lastRow).ClearFormats
Columns("A:A").Select
Selection.Copy
Columns("B:B").Select
Selection.PasteSpecial Paste:=xlPasteFormats
Application.CutCopyMode = False
Columns("A:A").ColumnWidth = 18.5
Columns("B:B").ColumnWidth = 18.5

Range("A1:B1").Select
Selection.AutoFilter
ActiveSheet.Range("$A$1:$B$10").AutoFilter Field:=1, Criteria1:="<>"

'Format the Adverse Event Status Report Sheet
WS1.Activate

ActiveSheet.Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.WrapText = True
Selection.VerticalAlignment = xlCenter
Selection.HorizontalAlignment = xlLeft
Selection.Borders.LineStyle = xlContinuous
Selection.Borders.Weight = xlThin
Selection.AutoFilter
ActiveSheet.Range("A1").AutoFilter Field:=FSField, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"

Application.ScreenUpdating = True
    
End Sub

Sub FormatCRFs()

Dim WB1 As Workbook
Dim WS1 As Worksheet
Dim WS2 As Worksheet
Dim lastRow As Long

Application.ScreenUpdating = False
Set WB1 = ActiveWorkbook
Set WS1 = ActiveSheet
'ActiveSheet.Name = "Raw Data"
'Count number of rows w/ header
WS1.Range("A1").Select
lastRow = ActiveSheet.Cells.Find(What:="*", SearchDirection:=xlPrevious).Row

Cells.Select
Selection.Copy

'Add WS2 as output
'With WB1
 '   .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Filtered Report"
 '   Set WS2 = Sheets("Filtered Report")
'End With

   ' Cells.Select
   ' ActiveSheet.Paste



If WS1.Range("B5").Value = "Basic Stats" Then 'report without header
    'delete first 7 rows and last 2 rows.
    Rows("1:7").Select
    Selection.Delete Shift:=xlUp
    Range("A" & (lastRow - 6) & ":A" & (lastRow - 5)).Select 'last row was returned before deletion of first 7 rows
ElseIf WS1.Range("B8").Value = "Basic Stats" Then 'report with header
'delete first 10 rows and last 2 rows.
    Rows("1:10").Select
    Selection.Delete Shift:=xlUp
    Range("A" & (lastRow - 9) & ":A" & (lastRow - 8)).Select 'last row was returned before deletion of first 10 rows
End If
Selection.EntireRow.Delete

'Format the completed sheet,hide the unneeded columns
'WS2.Activate
Call Main.OutFormat
Application.ScreenUpdating = True

End Sub


Sub FormatAE()
    Call MainAE
    Dim fileSaveName As Variant
    Dim modifiedDate As String
    modifiedDate = Now2Date(Now)
    Dim modifiedTime As String
    modifiedTime = Now2Time(Now)
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:=modifiedDate & "- XXXXX Adverse Events Report " & modifiedTime & " EST.xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    If fileSaveName = False Then
        MsgBox "You haven't saved the document", vbExclamation
    Else
        ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
    End If
End Sub

Sub FormatPDAE()
    Call MainAE
    Dim fileSaveName As Variant
    Dim modifiedDate As String
    modifiedDate = Now2Date(Now)
    Dim modifiedTime As String
    modifiedTime = Now2Time(Now)
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:=modifiedDate & "- XXXXX Protocol Defined Adverse Events Report " & modifiedTime & " EST.xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    If fileSaveName = False Then
        MsgBox "You haven't saved the document", vbExclamation
    Else
        ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
    End If
End Sub
