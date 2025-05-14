Attribute VB_Name = "ConmedAE"
Option Explicit

Sub MainConmed()

'This sub is shared between AdHoc and Quick ConMed reports

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

lastRow = FindLastRowA(WS1)
ActiveSheet.Name = "Conmed Status"

FSField = Application.Match("Form Status", WS1.Range("A1:AF1"), 0)
Set WS2 = Sheets.Add
WS2.Name = "Conmed Overview"
FSFL = Split(Cells(1, FSField).Address, "$")(1) 'split ($G$1","$")return the address of row 1 and column form status exisits (column G)
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
    ActiveSheet.Range("A1").Select
    FormatTable 'formatting WS2
    
'Format the Conmed Status Report Sheet
    
    WS1.Activate
    
    ActiveSheet.Range("A1").Select
   
    FormatTable 'formatting WS1
    
    Selection.AutoFilter
    ActiveSheet.Range("A1").AutoFilter Field:=FSField, Criteria1:="Incomplete", Operator:=xlOr, Criteria2:="Work In Progress"

    'Enable Screen Update
    Application.ScreenUpdating = True

End Sub

Sub MainAE()

'This sub is shared between AdHoc and Quick AE and PDAE reports

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
lastRow = FindLastRowA(WS1)

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
ActiveSheet.Range("A1").Select
FormatTable 'formatting WS2

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

'Enable Screen Update
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

Sub AdHocConMed()
'Format Ad Hoc Conmed Report
    'Removing unnecessary headers and rows
    Call FormatCRFs
    
    'Call the similar part of ad-hoc and Quick reports
    Call MainConmed
    
    'Prompt the user to update the file name and save the file
    Dim fileSaveName As Variant
    Dim modifiedDate As String
    modifiedDate = Now2Date(Now)
    Dim modifiedTime As String
    modifiedTime = Now2Time(Now)
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:=modifiedDate & "-XXXXX Concomitant Medications Report " & modifiedTime & " EST.xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    If fileSaveName = False Then
        MsgBox "You haven't saved the document", vbExclamation
    Else
        ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
    End If
    
End Sub

Sub AdHocAE()

'Format Ad Hoc AE Report
    'Removing unnecessary headers and rows
    Call FormatCRFs
    
    Call MainAE
    
    'Prompt the user to update the file name and save the file
    
    Dim fileSaveName As Variant
    Dim modifiedDate As String
    modifiedDate = Now2Date(Now)
    Dim modifiedTime As String
    modifiedTime = Now2Time(Now)
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:=modifiedDate & "-XXXXX Adverse Events Report " & modifiedTime & " EST.xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    If fileSaveName = False Then
        MsgBox "You haven't saved the document", vbExclamation
    Else
        ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
    End If
    
End Sub

Sub AdHocPDAE()

'Format Ad Hoc PDAE Report
    'Removing unnecessary headers and rows
    Call FormatCRFs
    
    Call MainAE
    
    'Prompt the user to update the file name and save the file
    
    Dim fileSaveName As Variant
    Dim modifiedDate As String
    modifiedDate = Now2Date(Now)
    Dim modifiedTime As String
    modifiedTime = Now2Time(Now)
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:=modifiedDate & "-XXXXX Protocol Defined Adverse Events Report " & modifiedTime & " EST.xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    If fileSaveName = False Then
        MsgBox "You haven't saved the document", vbExclamation
    Else
        ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
    End If
    
End Sub

Sub QuickConMed()
    
    'Formatting before calling MainConmed
    Call QuickRepCleanup 'Remove columns specified in the helper sub
    ActiveSheet.Rows(FindLastRowA(ActiveSheet)).Delete
    
    Call MainConmed 'Call MainConmed
    
    'Prompt the user to update the file name and save the file
    
    Dim fileSaveName As Variant
    Dim modifiedDate As String
    modifiedDate = Now2Date(Now)
    Dim modifiedTime As String
    modifiedTime = Now2Time(Now)
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:=modifiedDate & "-XXXXX Concomitant Medications Report " & modifiedTime & " EST.xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    If fileSaveName = False Then
        MsgBox "You haven't saved the document", vbExclamation
    Else
        ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
    End If
End Sub

Sub QuickAE()

    'Format Quick AE Report

    Call QuickRepCleanup 'Remove columns specified in the helper sub
    ActiveSheet.Rows(FindLastRowA(ActiveSheet)).Delete
    Call MainAE 'Call MainAE
    
    'Prompt the user to update the file name and save the file
    
    Dim fileSaveName As Variant
    Dim modifiedDate As String
    modifiedDate = Now2Date(Now)
    Dim modifiedTime As String
    modifiedTime = Now2Time(Now)
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:=modifiedDate & "-XXXXX Adverse Events Report " & modifiedTime & " EST.xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    If fileSaveName = False Then
        MsgBox "You haven't saved the document", vbExclamation
    Else
        ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
    End If
End Sub

Sub QuickPDAE()
  
    'Format Quick AE Report

    Call QuickRepCleanup 'Remove columns specified in the helper sub
    ActiveSheet.Rows(FindLastRowA(ActiveSheet)).Delete
    
    Call MainAE 'Call MainAE
    
    'Prompt the user to update the file name and save the file
    
    Dim fileSaveName As Variant
    Dim modifiedDate As String
    modifiedDate = Now2Date(Now)
    Dim modifiedTime As String
    modifiedTime = Now2Time(Now)
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:=modifiedDate & "-XXXXX Protocol Defined Adverse Events Report " & modifiedTime & " EST.xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    If fileSaveName = False Then
        MsgBox "You haven't saved the document", vbExclamation
    Else
        ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
    End If
End Sub
Sub NewQuickAE()
    Dim WB As Workbook
    Dim WSsrc As Worksheet
    Dim WSdest As Worksheet
    Dim lastRow As Long, lastColSrc As Long, lastColDest As Long
    Dim studyCode As String, rawID As String, parts() As String
    Dim headers As Variant, hdrIdx As Variant
    Dim i As Long, j As Long
    
    Set WB = ActiveWorkbook
    Set WSsrc = ActiveSheet
    
    Application.ScreenUpdating = False
    
    ' 0) Cleanup & delete last data row
    Call QuickRepCleanup
    lastRow = FindLastRowA(WSsrc)
    WSsrc.Rows(lastRow).Delete
    
    ' 1) Grab study code from A2
    rawID = Trim(WSsrc.Range("A2").Value)
    If Len(rawID) < 1 Then
        MsgBox "Cell A2 is empty; cannot determine study code.", vbExclamation
        Exit Sub
    End If
    
    studyCode = rawID
    If Len(studyCode) = 0 Then
        MsgBox "Unable to extract study code from A2.", vbExclamation
        Exit Sub
    End If
    
    ' 2) Get your ordered header list
    headers = GetColumnOrder(studyCode)
    If Not IsArray(headers) Or UBound(headers) < LBound(headers) Then
        ' fallback to MainAE if no custom order
        Call MainAE
        Exit Sub
    End If
    
    ' 3) Add a new sheet for the Quick AE Report
    With WB
        .Sheets.Add After:=.Sheets(.Sheets.Count)
        ActiveSheet.Name = "Quick AE Report"
        Set WSdest = ActiveSheet
    End With
    
    ' 4) Copy in the prioritized columns + headers
    Call CopyColumnsWithHeaders(WSsrc, WSdest, headers, 1, 1)
    

    
    ' 6) Copy every other column
    lastColSrc = FindLastColumn(WSsrc)
    lastColDest = FindLastColumn(WSdest)
    j = 0
    For i = 1 To lastColSrc
        If Not IsNumberInArray(i, hdrIdx) Then
            Call CopyColumnsWithIndex(WSsrc, WSdest, i, i + lastColDest - j)
        Else
            j = j + 1
        End If
    Next i
    
    ' 7) Apply your MainAE-style formatting to the new sheet
    WSdest.Activate
    Call FormatTable
    
    Application.ScreenUpdating = True
    
    ' 8) Prompt the user to save
    Dim fileSaveName As Variant
    fileSaveName = Application.GetSaveAsFilename( _
        InitialFileName:=Now2Date(Now) & " " & studyCode & " Safety Quick Report " & Now2Time(Now) & " EST.xlsx", _
        FileFilter:="Excel Files (*.xlsx), *.xlsx")
    If fileSaveName = False Then
        MsgBox "You haven't saved the document", vbExclamation
    Else
        WB.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
    End If
End Sub
Function GetColumnOrder(studyCode As String) As Variant
    Select Case studyCode
    Case "50424"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "Attribution to T-cell Therapy (IP1)", "T-cell Therapy Expectedness (IP1)", _
          "Other Attribution", "Specify Other Attribution", "Attribution to Ruxolitinib (IP2)", _
          "CTCAE Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Event onset", "Additional Toxicity Details", "Event ongoing?", _
          "Date the event became an SAE", "Seriousness as listed on the SAE form" _
        )
    Case "15CT055- CCI Adverse Event"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "T-cell Attribution", "T-cell Expectedness", _
          "Other, Specify", "Other Attribution", _
          "CTCAE Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Event onset", "Additional Toxicity Details", "Event ongoing?", _
          "Date the event became an SAE", "Action Taken", "Seriousness as listed on the SAE form" _
        )
    Case "19CT011-CCI Adverse Event"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "T-cell Attribution", "T-cell Expectedness", _
          "Specify", "Other Attribution", _
          "CTCAE Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Event onset", "Additional Toxicity Details", "Event ongoing?", _
          "Date the event became an SAE", "Action Taken", "Seriousness as listed on the SAE form" _
        )
    Case "01422-CCI Adverse Event"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "T-cell Therapy", "Tadekinig Alfa Attribution", "Tadekinig Alfa Expectedness", _
          "Specify", "Other Attribution", _
          "CTCAE Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Event onset", "Additional Toxicity Details", "Event ongoing?", _
          "Date the event became an SAE", "Action Taken", "Seriousness as listed on the SAE form" _
        )
    Case "01817- CCI Protocol Defined Adverse Events"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "Attribution", "Expectedness", _
          "Other, Specify", "Other Attribution", _
          "CTCAE Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Was event ongoing at study discontinuation?", "Additional Toxicity Details", _
          "Action Taken", "Date the event became an SAE:", "Seriousness as listed on the SAE form" _
        )
    Case "12320-CCI Adverse Event"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "T-cell Attribution", "T-cell Expectedness", _
          "Specify", "Other Attribution", _
          "CTCAE Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Event onset", "Additional Toxicity Details", "Event ongoing?", _
          "Date the event became an SAE", "Action Taken", "Seriousness as listed on the SAE form" _
        )
    Case "12418- CCI Adverse Event"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "T-cell Attribution", "T-cell Expectedness", _
          "Other, Specify", "Other Attribution", _
          "CTCAE Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Event onset", "Additional Toxicity Details", "Event ongoing?", _
          "Date the event became an SAE", "Action Taken", "Seriousness as listed on the SAE form" _
        )
    Case "12418- CCI PDAE"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "T-cell Attribution", "T-cell Expectedness", _
          "Other, Specify", "Other Attribution", _
          "CTCAE Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Additional Toxicity Details", "Event ongoing?", _
          "Date the event became an SAE", "Action Taken" _
        )
    Case "14217-CCI Adverse Event V2"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "T-cell Attribution", "T-cell Expectedness", _
          "Other, Specify", "Other Attribution", _
          "CTCAE Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Event onset", "Additional Toxicity Details", "Event ongoing?", _
          "Date the event became an SAE", "Seriousness as listed on the SAE form" _
        )
    Case "14217-CCI Protocol-Defined Adverse Event"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "T-cell Attribution", "T-cell Expectedness", _
          "Specify Other Attribution", "Other Attribution", _
          "CTCAE Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Additional Toxicity Details", "Event ongoing?", _
          "Date the event became an SAE", "Seriousness as listed on the SAE form" _
        )
    Case "19422-CCI Adverse Event"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "T-cell Attribution", "T-cell Expectedness", _
          "Specify Other Attribution", "Other Attribution", _
          "CTCAE Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Event onset", "Additional Toxicity Details", "Event ongoing?", _
          "Date the event became an SAE", "Seriousness as listed on the SAE form" _
        )
    Case "19422-CCI Protocol-Defined Adverse Event"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "T-cell Attribution", "T-cell Expectedness", _
          "Specify Other Attribution", "Other Attribution", _
          "CTCAE Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Additional Toxicity Details", "Event ongoing?", _
          "Date the event became an SAE", "Seriousness as listed on the SAE form" _
        )
    Case "35418-CCI Adverse Event"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "T-cell Attribution", "T-cell Expectedness", _
          "Other, Specify", "Other Attribution", _
          "CTCAE Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Event onset", "Additional Toxicity Details", "Event ongoing?", _
          "Date the event became an SAE", "Action Taken", "Seriousness as listed on the SAE form" _
        )
    Case "46417-CCI Adverse Events"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "T-cell Attribution", "T-cell Expectedness", _
          "Other, Specify", "Other Attribution", _
          "CTCAE Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Event onset", "Additional Toxicity Details", "Was event ongoing at study discontinuation?", _
          "Date the event became an SAE", "Seriousness as listed on the SAE form" _
        )
    Case "46417-CCI Protocol Defined Adverse Events"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "T-cell Attribution", "T-cell Expectedness", _
          "Other, Specify", "Other Attribution", _
          "CTCAE Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Additional Toxicity Details", "Was event ongoing at study discontinuation?", _
          "Date the event became an SAE", "Seriousness as listed on the SAE form" _
        )
    Case "CD4CAR-ZFN-CCI Adverse Event V2"
        GetColumnOrder = Array( _
          "Subject ID#", "AE or SAE?", "T-cell Attribution", "T-cell Expectedness", _
          "Other, Specify", "Other Attribution", _
          "Category", "Toxicity", "Grade", "Start Date", "Stop Date", _
          "Event onset", "Event ongoing?", _
          "Date the event became an SAE", "Seriousness as listed on the SAE form" _
        )
    Case Else
        GetColumnOrder = Array()  ' no custom order defined
    End Select

End Function
