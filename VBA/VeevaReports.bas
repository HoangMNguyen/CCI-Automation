Attribute VB_Name = "VeevaReports"
Sub SubjectProgressListingFormat()
    Application.ScreenUpdating = False
    Dim StudyNum As String
    StudyNum = GetStudyNumber()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(1)
    ws.Name = StudyNum + "_Subject_Progress"
    
    ' Remove columns by calling RemoveColumn function
    Call RemoveColumn(ws, "Study")
    Call RemoveColumn(ws, "Country")
    Call RemoveColumn(ws, "Site")
    Call RemoveColumn(ws, "Forms Not Frozen")
    Call RemoveColumn(ws, "Event Dates Not Frozen")
    Call RemoveColumn(ws, "Forms Not Locked")
    Call RemoveColumn(ws, "Event Dates Not Locked")
    Call RemoveColumn(ws, "Forms Not Signed")
    Call RemoveColumn(ws, "Event Dates Not Signed")
    Call RemoveColumn(ws, "Total MC")
    Call RemoveColumn(ws, "MedDRA MC")
    Call RemoveColumn(ws, "WHODrug MC")
    Call RemoveColumn(ws, "MedDRA MC Need Coding")
    Call RemoveColumn(ws, "WHODrug MC Need Coding")
    Call RemoveColumn(ws, "Subject Vault ID")
    
    Call OutFormat
    Application.ScreenUpdating = True
    
    'Save file
    Dim modifiedDate As String
    Dim modifiedTime As String
    Dim fileSaveName As Variant
    modifiedDate = Now2Date(Now)
    modifiedTime = Now2Time(Now)
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:=modifiedDate & "-" & StudyNum & " Subject Progress Report " & modifiedTime & " EST.xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    
    If fileSaveName = False Then
        MsgBox "You haven't saved the document", vbExclamation
    Else
        ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
    End If

End Sub

Sub QueryDetailListingFormatSponsor()
    Application.ScreenUpdating = False
    Dim WS1 As Worksheet
    Dim StudyNum As String
    Set WS1 = ActiveWorkbook.Sheets(1)
    StudyNum = GetStudyNumber()
    WS1.Name = StudyNum + "_Query_Detail"
    ' Remove columns by calling RemoveColumn function
    Call RemoveColumn(WS1, "Study")
    Call RemoveColumn(WS1, "Country")
    Call RemoveColumn(WS1, "Site")
    Call RemoveColumn(WS1, "Event Label")
    Call RemoveColumn(WS1, "Query ID")
    Call RemoveColumn(WS1, "Query Rule")
    Call RemoveColumn(WS1, "Created By Role")
    Call RemoveColumn(WS1, "Created By Query Team")
    Call RemoveColumn(WS1, "Answered By Role")
    Call RemoveColumn(WS1, "Closed By Role")
    Call RemoveColumn(WS1, "Query Vault ID")
    Call RemoveColumn(WS1, "Event Group Sequence Number")
    Call RemoveColumn(WS1, "Item OID")
    Call RemoveColumn(WS1, "Query Team")
    Call RemoveColumn(WS1, "Answered By Query Team")
    Call RemoveColumn(WS1, "Closed By Query Team")
    
    Call OutFormat
    
    Call SetWidthWrapColumn(WS1, "Item Label", 60)
    Call SetWidthWrapColumn(WS1, "Original Query Text", 60)
    Call SetWidthWrapColumn(WS1, "Latest Query Comment", 60)
    Call SetWidthWrapColumn(WS1, "Latest Query Answer Text", 60)
    Call SetWidthWrapColumn(WS1, "Item Value Now", 60)
    Call SetWidthWrapColumn(WS1, "Item Value Before Query", 60)
    
    'Round down column
    Call RoundDownColumn(WS1, "Days Unresolved")
    
    
    Dim WS2 As Worksheet
    Set WS2 = Sheets.Add(After:=WS1)
    WS2.Name = "Query Status Summary"
    WS2.Range("A1").Value = "Query Status"
    WS2.Range("B1").Value = "Counts"
    WS2.Range("A2").Value = "Answered"
    WS2.Range("B2").Value = CountPerColumnName(WS1, "Query Status", "Answered")
    WS2.Range("A3").Value = "Closed"
    WS2.Range("B3").Value = CountPerColumnName(WS1, "Query Status", "Closed")
    WS2.Range("A4").Value = "Open"
    WS2.Range("B4").Value = CountPerColumnName(WS1, "Query Status", "Open")
    
    WS2.Range("A6").Value = "Open Query Form Types"
    WS2.Range("B6").Value = "Counts"
    WS2.Range("A7").Value = "Adverse Event"
    WS2.Range("B7").Value = CountPerColumnName(WS1, "Query Status", "Open", "Form Label", "Adverse Event")
    WS2.Range("A8").Value = "Concomitant Medications"
    WS2.Range("B8").Value = CountPerColumnName(WS1, "Query Status", "Open", "Form Label", "Concomitant Medication")
    WS2.Range("A9").Value = "LTFU ANP"
    WS2.Range("B9").Value = CountPerColumnName(WS1, "Query Status", "Open", "Form Label", "LTFU Antineoplastic (ANP) Therapy")
    WS2.Range("A10").Value = "All Other Forms"
    WS2.Range("B10").Value = CountPerColumnName(WS1, "Query Status", "Open") - WS2.Range("B7").Value - WS2.Range("B8").Value - WS2.Range("B9").Value
    
    WS2.Activate
    ActiveSheet.Range("A1").Select
    Call FormatTable
    
    ActiveSheet.Range("A6").Select
    Call FormatTable
    
    'Filter values
    Dim FilterValues() As Variant
    FilterValues = Array("Open", "Answered")
    Call FilterColumn(WS1, "Query Status", FilterValues)
    
    Application.ScreenUpdating = True
    
    'Save file
    Dim modifiedDate As String
    Dim modifiedTime As String
    Dim fileSaveName As Variant
    modifiedDate = Now2Date(Now)
    modifiedTime = Now2Time(Now)
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:=modifiedDate & "-" & StudyNum & " Query Report " & modifiedTime & " EST_Sponsor" & ".xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    
    If fileSaveName = False Then
        MsgBox "You haven't saved the document", vbExclamation
    Else
        ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
    End If

End Sub

Sub QueryDetailListingFormatSite()
    Application.ScreenUpdating = False
    Dim WS1 As Worksheet
    Dim StudyNum As String
    Set WS1 = ActiveWorkbook.Sheets(1)
    StudyNum = GetStudyNumber()
    WS1.Name = StudyNum + "_Query_Detail"
    ' Remove columns by calling RemoveColumn function
    Call RemoveColumn(WS1, "Study")
    Call RemoveColumn(WS1, "Country")
    Call RemoveColumn(WS1, "Site")
    Call RemoveColumn(WS1, "Event Label")
    Call RemoveColumn(WS1, "Query ID")
    Call RemoveColumn(WS1, "Query Rule")
    Call RemoveColumn(WS1, "Created By Role")
    Call RemoveColumn(WS1, "Created By Query Team")
    Call RemoveColumn(WS1, "Answered By Role")
    Call RemoveColumn(WS1, "Closed By Role")
    Call RemoveColumn(WS1, "Query Vault ID")
    Call RemoveColumn(WS1, "Event Group Sequence Number")
    Call RemoveColumn(WS1, "Item OID")
    Call RemoveColumn(WS1, "Query Team")
    Call RemoveColumn(WS1, "Answered By Query Team")
    Call RemoveColumn(WS1, "Closed By Query Team")
    Call RemoveColumn(WS1, "Closed Date")
    Call RemoveColumn(WS1, "Closed By")
    
    Call OutFormat
    
    Call SetWidthWrapColumn(WS1, "Item Label", 60)
    Call SetWidthWrapColumn(WS1, "Original Query Text", 60)
    Call SetWidthWrapColumn(WS1, "Latest Query Comment", 60)
    Call SetWidthWrapColumn(WS1, "Latest Query Answer Text", 60)
    Call SetWidthWrapColumn(WS1, "Item Value Now", 60)
    Call SetWidthWrapColumn(WS1, "Item Value Before Query", 60)
    
    'Round down column
    Call RoundDownColumn(WS1, "Days Unresolved")

    
    'Filter values
    Dim FilterValues() As Variant
    FilterValues = Array("Open")
    Call FilterColumn(WS1, "Query Status", FilterValues)
    
    Application.ScreenUpdating = True
    
    'Save file
    Dim modifiedDate As String
    Dim modifiedTime As String
    Dim fileSaveName As Variant
    modifiedDate = Now2Date(Now)
    modifiedTime = Now2Time(Now)
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:=modifiedDate & "-" & StudyNum & " Query Report " & modifiedTime & " EST_Site" & ".xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    
    If fileSaveName = False Then
        MsgBox "You haven't saved the document", vbExclamation
    Else
        ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
    End If

End Sub

Sub FormProgressListingFormat()
    Application.ScreenUpdating = False
    
    Dim WS1 As Worksheet
    Set WS1 = ActiveWorkbook.Sheets(1)
    Dim StudyNum As String
    StudyNum = GetStudyNumber()
    WS1.Name = StudyNum + "_Form_Progress"
    
    ' Remove columns by calling RemoveColumn function
    Call RemoveColumn(WS1, "Study")
    Call RemoveColumn(WS1, "Country")
    Call RemoveColumn(WS1, "Site")
    Call RemoveColumn(WS1, "Event Label")
    Call RemoveColumn(WS1, "Event Group Sequence Number")
    Call RemoveColumn(WS1, "SDV Plan")
    Call RemoveColumn(WS1, "SDV Override Plan")
    Call RemoveColumn(WS1, "SDV Required")
    Call RemoveColumn(WS1, "DMR Plan")
    Call RemoveColumn(WS1, "DMR Override Plan")
    Call RemoveColumn(WS1, "DMR Required")
    Call RemoveColumn(WS1, "Frozen")
    Call RemoveColumn(WS1, "Freeze Date")
    Call RemoveColumn(WS1, "Locked")
    Call RemoveColumn(WS1, "Lock Date")
    Call RemoveColumn(WS1, "Signed")
    Call RemoveColumn(WS1, "Sign Date")
    Call RemoveColumn(WS1, "Last Signed Date")
    Call RemoveColumn(WS1, "Form Vault ID")
    
    

    Dim WS2 As Worksheet
    Set WS2 = Sheets.Add(After:=WS1)
    WS2.Name = "Forms Progress Metrics"
    WS2.Range("A1").Value = "Form Status"
    WS2.Range("B1").Value = "Counts"
    WS2.Range("A2").Value = "Blank"
    WS2.Range("B2").Value = CountPerColumnName(WS1, "Form Status", "Blank")
    WS2.Range("A3").Value = "In Progress"
    WS2.Range("B3").Value = CountPerColumnName(WS1, "Form Status", "In Progress")
    WS2.Range("A4").Value = "In Progress Post Submit"
    WS2.Range("B4").Value = CountPerColumnName(WS1, "Form Status", "In Progress Post Submit")
    WS2.Range("A5").Value = "Submitted"
    WS2.Range("B5").Value = CountPerColumnName(WS1, "Form Status", "Submitted")
    
    WS2.Range("A7").Value = "SDV Status"
    WS2.Range("B7").Value = "Counts"
    WS2.Range("A8").Value = "Forms Pending SDV"
    WS2.Range("B8").Value = CountPerColumnName(WS1, "Form Status", "Submitted", "SDV Complete", "No")
    WS2.Range("A9").Value = "Forms SDV Complete"
    WS2.Range("B9").Value = CountPerColumnName(WS1, "Form Status", "Submitted", "SDV Complete", "Yes")
    WS2.Range("A10").Value = "% Forms SDV Complete"
    WS2.Range("B10").Value = Int(CountPerColumnName(WS1, "Form Status", "Submitted", "SDV Complete", "Yes") / (CountPerColumnName(WS1, "Form Status", "Submitted", "SDV Complete", "No") + CountPerColumnName(WS1, "Form Status", "Submitted", "SDV Complete", "Yes")) * 100)
    
    WS2.Range("A13").Value = "DMR Status"
    WS2.Range("B13").Value = "Counts"
    WS2.Range("A14").Value = "Forms Pending DMR"
    WS2.Range("B14").Value = CountPerColumnName(WS1, "Form Status", "Submitted", "DMR Complete", "No")
    WS2.Range("A15").Value = "Forms DMR Complete"
    WS2.Range("B15").Value = CountPerColumnName(WS1, "Form Status", "Submitted", "DMR Complete", "Yes")
    WS2.Range("A16").Value = "% Forms DMR Complete"
    WS2.Range("B16").Value = Int(CountPerColumnName(WS1, "Form Status", "Submitted", "DMR Complete", "Yes") / (CountPerColumnName(WS1, "Form Status", "Submitted", "DMR Complete", "No") + CountPerColumnName(WS1, "Form Status", "Submitted", "DMR Complete", "Yes")) * 100)
    
    'Adding if ReSDV or ReDMR exists
    Dim ReSDVCountYes As Integer
    Dim ReDMRCountYes As Integer
    Dim ReSDVCountNo As Integer
    Dim ReDMRCountNo As Integer
    ReSDVCountYes = CountPerColumnName(WS1, "Form Status", "Submitted", "Requires Re-SDV", "Yes")
    ReDMRCountYes = CountPerColumnName(WS1, "Form Status", "Submitted", "Requires Re-DMR", "Yes")
    ReSDVCountNo = CountPerColumnName(WS1, "Form Status", "Submitted", "Requires Re-SDV", "No")
    ReDMRCountNo = CountPerColumnName(WS1, "Form Status", "Submitted", "Requires Re-DMR", "No")
    
    If ReSDVCountYes <> 0 Or ReDMRCountYes <> 0 Or ReSDVCountNo <> 0 Or ReDMRCountNo <> 0 Then
        WS2.Range("A11").Value = "Forms Require Re-SDV"
        WS2.Range("B11").Value = ReSDVCountYes
        
        WS2.Range("A17").Value = "Forms Require Re-DMR"
        WS2.Range("B17").Value = ReDMRCountYes
    End If
    
    
    'Errors tab
    WS2.Range("D1").Value = "Error Description"
    WS2.Range("E1").Value = "Subject"
    WS2.Range("F1").Value = "Subject Status"
    WS2.Range("G1").Value = "Event Group Label"
    WS2.Range("H1").Value = "Event Date"
    WS2.Range("I1").Value = "Form Label"
    WS2.Range("J1").Value = "Form Sequence Number"
    WS2.Range("K1").Value = "Form Status"
    WS2.Range("L1").Value = "SDV Complete"
    WS2.Range("M1").Value = "DMR Complete"
    WS2.Range("N1").Value = "Open Queries"
    
    
    Dim NoSDVYesDMR As Integer
    

    
    NoSDVYesDMR = CountPerColumnName(WS1, "SDV Complete", "No", "DMR Complete", "Yes")
    
    WS1.Activate
    Call RemoveFilter
    Call FilterColumn(WS1, "SDV Complete", "No")
    Call FilterColumn(WS1, "DMR Complete", "Yes")
    NoSDVYesDMR = CountFilteredRows(WS1, "Subject")
    Dim HeadersToCopy As Variant
    HeadersToCopy = Array("Subject", "Subject Status", "Event Group Label", "Event Date", "Form Label", "Form Sequence Number", "Form Status", "SDV Complete", "DMR Complete", "Open Queries")  ' Replace with your headers
    Call CopySelectedVisibleColumnsToLocation(WS1, WS2, HeadersToCopy, 5, 2)
    
    Dim StartingRow As Integer
    StartingRow = 1
    If NoSDVYesDMR > 0 Then
        For i = 1 To NoSDVYesDMR
            WS2.Range("D" & (i + StartingRow)).Value = "SDV is No but DMR is Yes"
        Next i
    End If
    StartingRow = StartingRow + NoSDVYesDMR

    Dim YesSDVOpenQuery As Integer
    Call RemoveFilter
    Call FilterColumn(WS1, "SDV Complete", "Yes")
    Call FilterExcludeValue(WS1, "Open Queries", "0")
    YesSDVOpenQuery = CountFilteredRows(WS1, "Subject")
    Call CopySelectedVisibleColumnsToLocation(WS1, WS2, HeadersToCopy, 5, StartingRow + 1)
    If YesSDVOpenQuery > 0 Then
        For i = 1 To YesSDVOpenQuery
            WS2.Range("D" & (i + StartingRow)).Value = "SDV is Yes but there are open queries"
        Next i
    End If
    StartingRow = StartingRow + YesSDVOpenQuery
    
    
    Dim YesDMROpenQuery As Integer
    Call RemoveFilter
    Call FilterColumn(WS1, "DMR Complete", "Yes")
    Call FilterExcludeValue(WS1, "Open Queries", "0")
    YesDMROpenQuery = CountFilteredRows(WS1, "Subject")
    Call CopySelectedVisibleColumnsToLocation(WS1, WS2, HeadersToCopy, 5, StartingRow + 1)
    If YesDMROpenQuery > 0 Then
        For i = 1 To YesDMROpenQuery
            WS2.Range("D" & (i + StartingRow)).Value = "DMR is Yes but there are open queries"
        Next i
    End If
    StartingRow = StartingRow + YesDMROpenQuery
    
    'Not Submitted but SDV Yes
    Dim NotSubmittedYesSDV As Integer
    Call RemoveFilter
    Call FilterExcludeValue(WS1, "Form Status", "Submitted")
    Call FilterColumn(WS1, "SDV Complete", "Yes")
    NotSubmittedYesSDV = CountFilteredRows(WS1, "Subject")
    Call CopySelectedVisibleColumnsToLocation(WS1, WS2, HeadersToCopy, 5, StartingRow + 1)
    If NotSubmittedYesSDV > 0 Then
        For i = 1 To NotSubmittedYesSDV
            WS2.Range("D" & (i + StartingRow)).Value = "Form is not Submitted but SDV is Yes"
        Next i
    End If
    StartingRow = StartingRow + NotSubmittedYesSDV
    
    'Not Submitted but DMR Yes
    Dim NotSubmittedYesDMR As Integer
    Call RemoveFilter
    Call FilterExcludeValue(WS1, "Form Status", "Submitted")
    Call FilterColumn(WS1, "DMR Complete", "Yes")
    NotSubmittedYesDMR = CountFilteredRows(WS1, "Subject")

    Call CopySelectedVisibleColumnsToLocation(WS1, WS2, HeadersToCopy, 5, StartingRow + 1)
    If NotSubmittedYesDMR > 0 Then
        For i = 1 To NotSubmittedYesDMR
            WS2.Range("D" & (i + StartingRow)).Value = "Form is not Submitted but DMR is Yes"
        Next i
    End If
    StartingRow = StartingRow + NotSubmittedYesDMR
    
    
    
    
    'Formatting
    WS2.Activate
    ActiveSheet.Range("A1").Select
    Call FormatTable
    
    ActiveSheet.Range("A7").Select
    Call FormatTable
    
    ActiveSheet.Range("A13").Select
    Call FormatTable
    
    ActiveSheet.Range("D1").Select
    Call FormatTable
    
    WS1.Activate
    Call OutFormat
    
    'Filter values
    Dim FilterValues() As Variant
    FilterValues = Array("In Progress", "In Progress Post Submit", "Submitted")
    Call FilterColumn(WS1, "Form Status", FilterValues)
    
    Application.ScreenUpdating = True
    'Save file
    Dim modifiedDate As String
    Dim modifiedTime As String
    Dim fileSaveName As Variant
    modifiedDate = Now2Date(Now)
    modifiedTime = Now2Time(Now)
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:=modifiedDate & "-" & StudyNum & " Form Status Report " & modifiedTime & " EST.xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    
    If fileSaveName = False Then
        MsgBox "You haven't saved the document", vbExclamation
    Else
        ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
    End If
    

End Sub

