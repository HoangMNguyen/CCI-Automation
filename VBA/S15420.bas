Attribute VB_Name = "S15420"

Sub DSMB_Report()

Dim DMS As Worksheet, DMLastRow As Long
Dim IES As Worksheet, IELastRow As Long
Dim INFS As Worksheet, INFLastRow As Long
Dim DLAS As Worksheet, DLALastRow As Long
Dim AES As Worksheet, AELastRow As Long
Dim DIAGS As Worksheet, DIAGLastRow As Long
Dim RSS As Worksheet, RSLastRow As Long
Dim ITLFS As Worksheet, ITLFLastRow As Long
Dim ITRTS As Worksheet, ITRTLastRow As Long
Dim DSMBWB As Workbook, DSMBS As Worksheet

Dim dict As Scripting.Dictionary

Dim FilePath As String

Dim CL As Long

Set dict = New Scripting.Dictionary

CL = 1

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Set DSMBWB = Workbooks.Add
DSMBWB.Activate
DSMBWB.Sheets(1).Name = "DSMB Report"
Set DSMBS = DSMBWB.Sheets("DSMB Report")
'Add sheets to new workbook
FilePath = SelectFolder()
Call ImportSheetByName(FilePath, "15420_huCART19-IL18_AE")
Set AES = DSMBWB.Sheets("15420_huCART19-IL18_AE")
AELastRow = FindLastRowA(AES)
dict.Add "AE_LR", AELastRow

Call ImportSheetByName(FilePath, "15420_huCART19-IL18_DLA")
Set DLAS = DSMBWB.Sheets("15420_huCART19-IL18_DLA")
DLALastRow = FindLastRowA(DLAS)
dict.Add "DLA_LR", DLALastRow

Call ImportSheetByName(FilePath, "15420_huCART19-IL18_DM")
Set DMS = DSMBWB.Sheets("15420_huCART19-IL18_DM")
DMLastRow = FindLastRowA(DMS)
dict.Add "DM_LR", DMLastRow

Call ImportSheetByName(FilePath, "15420_huCART19-IL18_IE")
Set IES = DSMBWB.Sheets("15420_huCART19-IL18_IE")
IELastRow = FindLastRowA(IES)
dict.Add "IE_LR", IELastRow

Call ImportSheetByName(FilePath, "15420_huCART19-IL18_INF")
Set INFS = DSMBWB.Sheets("15420_huCART19-IL18_INF")
INFLastRow = FindLastRowA(INFS)
dict.Add "INF_LR", INFLastRow


Call ImportSheetByName(FilePath, "15420_huCART19-IL18_PRDIAG")
Set DIAGS = DSMBWB.Sheets("15420_huCART19-IL18_PRDIAG")
DIAGLastRow = FindLastRowA(DIAGS)
dict.Add "DIAG_LR", DIAGLastRow

Call ImportSheetByName(FilePath, "15420_huCART19-IL18_RS")
Set RSS = DSMBWB.Sheets("15420_huCART19-IL18_RS")
RSLastRow = FindLastRowA(RSS)
dict.Add "RS_LR", RSLastRow

Call ImportSheetByName(FilePath, "15420_huCART19-IL18_INITLF")
Set ITLFS = DSMBWB.Sheets("15420_huCART19-IL18_INITLF")
ITLFLastRow = FindLastRowA(ITLFS)
dict.Add "ITLF_LR", ITLFLastRow

Call ImportSheetByName(FilePath, "15420_huCART19-IL18_DSINITRT")
Set ITRTS = DSMBWB.Sheets("15420_huCART19-IL18_DSINITRT")
ITRTLastRow = FindLastRowA(ITRTS)
dict.Add "ITRT_LR", ITRTLastRow



'Add CL
dict.Add "CL", ColumnNum
'Adding values for each sheets to dict
dict.Add "DM_LSex", "Z"
dict.Add "DM_GI", "AA"
dict.Add "DM_SA", "AB"
dict.Add "DM_Age", "Y"
dict.Add "DM_Race", "AC"


dict.Add "INF_INF", "Y"
dict.Add "INF_SDay", "AL"
dict.Add "INF_Date", "AE"
dict.Add "INF_TCD1", "AF"
dict.Add "INF_TCD2", "AG"
dict.Add "INF_TCDA1", "AH"
dict.Add "INF_TCDA2", "AI"
dict.Add "INF_PFlow", "AK"

dict.Add "IE_SF", "AB"
dict.Add "IE_Enrolled", "AB"
dict.Add "IE_CSDate", "Z"

dict.Add "DLA_Dose", "Y"

dict.Add "DIAG_DT", "Y"
dict.Add "DIAG_PDCLL", "Z"
dict.Add "DIAG_PDCLL2", "AA"

dict.Add "DIAG_PDNHL", "AJ"
dict.Add "DIAG_PDNHL2", "AK"

dict.Add "AE_Attribution", "AG"
dict.Add "AE_CTCAE", "Y"
dict.Add "AE_Toxicity", "Z"
dict.Add "AE_Grade", "AB"
dict.Add "AE_StartDate", "AC"
dict.Add "AE_StopDate", "AE"


Call T1(DSMBWS, DMS, IES, INFS, DLAS, DIAGS, AES, RSS, ITLFS, DSMBS, dict)



Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub T1(DSMBWS, DMS, IES, INFS, DLAS, DIAGS, AES, RSS, ITLFS, DSMBS, dict)

Dim Screenfailed As Long
Dim Enrolled As Long
Dim Infused As Long
Dim MaleTS As Long
Dim BirthDay As Variant
Dim CSDate As Variant
Dim DraftLastRow As Long
Dim Draft As Worksheet

'STEP 1:
'Add values from dicts
LC = dict("LC")


'STEP 2:
'Add Draft sheet
ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(Sheets.Count)
Set Draft = ActiveWorkbook.Sheets(Sheets.Count)
ActiveWorkbook.Sheets(Sheets.Count).Name = "Draft"


'1)Preparing data in Draft Sheet
'Header
With Draft
    .Range("A1").Value = "Subject ID"
    .Range("B1").Value = "Dose Level"
    .Range("C1").Value = "Disease"
    .Range("D1").Value = "Legal Sex " & vbNewLine & "(Male/Female)"
    .Range("E1").Value = "Gender Identity"
    .Range("F1").Value = "Sex Assigned at Birth " & vbNewLine & "(Male/Female)"
    .Range("G1").Value = "Age"
    .Range("H1").Value = "Race"
    .Range("I1").Value = "Infused" & vbNewLine & "(Yes/No)"
    .Range("J1").Value = "Subject meets all study eligibility?"
    .Range("K1").Value = "Enrolled"
    .Range("L1").Value = "Study Day (Primary or Retreatment)"
End With

'Adding data
'Subject ID. CopyUnique also sort the subject ID with ascending order
Call CopyUnique(DMS, "D2:D" & dict("DM_LR"), Draft, "A2:A" & dict("DM_LR"))

DIAGS.Activate
ActiveSheet.Cells.Replace Chr(160), ""

Draft.Activate

On Error Resume Next
For i = 2 To dict("DM_LR")
    'Dose Level
    Cells(i, 2).Value = DLAS.Range(dict("DLA_Dose") & WorksheetFunction.Match(Cells(i, 1).Value, DLAS.Range("D1:D" & dict("DLA_LR")), 0))
    'Disease
    If Trim(DIAGS.Range(dict("DIAG_PDCLL") & WorksheetFunction.Match(Cells(i, 1).Value, DIAGS.Range("D1:D" & dict("DIAG_LR")), 0)).Value) = "" Then
        If Trim(DIAGS.Range(dict("DIAG_PDNHL2") & WorksheetFunction.Match(Cells(i, 1).Value, DIAGS.Range("D1:D" & dict("DIAG_LR")), 0)).Value) = "" Then
            Cells(i, 3).Value = DIAGS.Range(dict("DIAG_PDNHL") & WorksheetFunction.Match(Cells(i, 1).Value, DIAGS.Range("D1:D" & dict("DIAG_LR")), 0))
        Else
            Cells(i, 3).Value = DIAGS.Range(dict("DIAG_PDNHL2") & WorksheetFunction.Match(Cells(i, 1).Value, DIAGS.Range("D1:D" & dict("DIAG_LR")), 0))
        End If
    Else
        If DIAGS.Range(dict("DIAG_PDCLL2") & WorksheetFunction.Match(Cells(i, 1).Value, DIAGS.Range("D1:D" & dict("DIAG_LR")), 0)) = "" Then
            Cells(i, 3).Value = DIAGS.Range(dict("DIAG_PDCLL") & WorksheetFunction.Match(Cells(i, 1).Value, DIAGS.Range("D1:D" & dict("DIAG_LR")), 0))
        Else
            Cells(i, 3).Value = DIAGS.Range(dict("DIAG_PDCLL2") & WorksheetFunction.Match(Cells(i, 1).Value, DIAGS.Range("D1:D" & dict("DIAG_LR")), 0))
        End If
    End If
    'Legal Sex
    Cells(i, 4).Value = DMS.Range(dict("DM_LSex") & WorksheetFunction.Match(Cells(i, 1).Value, DMS.Range("D1:D" & dict("DM_LR")), 0))
    'Gender Identity
    Cells(i, 5).Value = DMS.Range(dict("DM_GI") & WorksheetFunction.Match(Cells(i, 1).Value, DMS.Range("D1:D" & dict("DM_LR")), 0))
    'Sex Assigned at birth
    Cells(i, 6).Value = DMS.Range(dict("DM_SA") & WorksheetFunction.Match(Cells(i, 1).Value, DMS.Range("D1:D" & dict("DM_LR")), 0))
    'Age
    BirthDay = DMS.Range(dict("DM_Age") & WorksheetFunction.Match(Cells(i, 1).Value, DMS.Range("D1:D" & dict("DM_LR")), 0))
    CSDate = IES.Range(dict("IE_CSDate") & WorksheetFunction.Match(Cells(i, 1).Value, IES.Range("D1:D" & dict("IE_LR")), 0))
    If VarType(CSDate) = 7 And VarType(BirthDay) = 7 Then
        Cells(i, 7).Value = Application.WorksheetFunction.RoundDown(WorksheetFunction.YearFrac(BirthDay, CSDate), 0)
    End If
    'Race (column AC of DM)
    Cells(i, 8).Value = DMS.Range(dict("DM_Race") & WorksheetFunction.Match(Cells(i, 1).Value, DMS.Range("D1:D" & dict("DM_LR")), 0))
    'Infused (column Y of INF)
    Cells(i, 9).Value = INFS.Range(dict("INF_INF") & WorksheetFunction.Match(Cells(i, 1).Value, INFS.Range("D1:D" & dict("INF_LR")), 0))
    If Cells(i, 9).Value <> "Yes" Then
        Cells(i, 9).Value = "No"
    End If
    'Screen Failed (column AB of IE)
    Cells(i, 10).Value = IES.Range(dict("IE_SF") & WorksheetFunction.Match(Cells(i, 1).Value, IES.Range("D1:D" & dict("IE_LR")), 0))
    'Enrolled (column AB of IE)
    Cells(i, 11).Value = IES.Range(dict("IE_Enrolled") & WorksheetFunction.Match(Cells(i, 1).Value, IES.Range("D1:D" & dict("IE_LR")), 0))
    'Study day (column AL of INF)
    Cells(i, 12).Value = INFS.Range(dict("INF_SDay") & WorksheetFunction.Match(Cells(i, 1).Value, INFS.Range("D1:D" & dict("INF_LR")), 0))
Next i
    
On Error GoTo 0

DraftLastRow = FindLastRowA(Draft)
dict.Add "Draft_LR", DraftLastRow

'//Final Result

'HEADER

'count number of total screened
DSMBS.Range("B1").Value = "Total Screened" & vbNewLine & "N=" & Str(dict("DM_LR") - 1)

'count number of screen failed
Screenfailed = CountPerColumn(Draft, 10, "No")
DSMBS.Range("C1").Value = "Screen Failed" & vbNewLine & "N=" & Str(Screenfailed)

'count number of Enrolled
Enrolled = CountPerColumn(Draft, 10, "Yes")
DSMBS.Range("D1").Value = "Enrolled" & vbNewLine & "N=" & Str(Enrolled)

'count number of Infused
Infused = CountPerColumn(Draft, 9, "Yes")
DSMBS.Range("E1").Value = "Infused" & vbNewLine & "N=" & Str(Infused)


With DSMBS
    .Range("A1").Value = "Status"
    .Range("A2").Value = "Sex"
    .Range("A2:E2").Merge
    .Range("A2:E2").HorizontalAlignment = xlCenter
    .Range("A3").Value = "Male"
    .Range("A4").Value = "Female"
    .Range("A5").Value = "Age"
    .Range("A5:E5").Merge
    .Range("A5:E5").HorizontalAlignment = xlCenter
    .Range("A6").Value = "Mean (SD)"
    .Range("A7").Value = "Median"
    .Range("A8").Value = "Range"
    .Range("A9").Value = "Race"
    .Range("A9:E9").Merge
    .Range("A9:E9").HorizontalAlignment = xlCenter
    .Range("A10").Value = "African American"
    .Range("A11").Value = "Alaska Native"
    .Range("A12").Value = "American Indian"
    .Range("A13").Value = "Asian"
    .Range("A14").Value = "Caucasian"
    .Range("A15").Value = "Multiple Races"
    .Range("A16").Value = "Pacific Islander"
    .Range("A17").Value = "Other"
    .Range("A18").Value = "Unknown"
End With



    'male and female
    On Error Resume Next
    With DSMBS
    .Range("B3").Value = Str(CountPerColumn(DMS, 26, "Male")) & " " & CountToPercent(CountPerColumn(DMS, 26, "Male"), dict("DM_LR") - 1)
    .Range("B4").Value = Str(CountPerColumn(DMS, 26, "Female")) & " " & CountToPercent(CountPerColumn(DMS, 26, "Female"), dict("DM_LR") - 1)
    .Range("C3").Value = Str(CountPerColumn(Draft, 4, "Male", 10, "No")) & " " & CountToPercent(CountPerColumn(Draft, 4, "Male", 10, "No"), CountPerColumn(Draft, 10, "No"))
    .Range("C4").Value = Str(CountPerColumn(Draft, 4, "Female", 10, "No")) & " " & CountToPercent(CountPerColumn(Draft, 4, "Female", 10, "No"), CountPerColumn(Draft, 10, "No"))
    .Range("D3").Value = Str(CountPerColumn(Draft, 4, "Male", 10, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 4, "Male", 10, "Yes"), CountPerColumn(Draft, 10, "Yes"))
    .Range("D4").Value = Str(CountPerColumn(Draft, 4, "Female", 10, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 4, "Female", 10, "Yes"), CountPerColumn(Draft, 10, "Yes"))
    .Range("E3").Value = Str(CountPerColumn(Draft, 4, "Male", 9, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 4, "Male", 9, "Yes"), CountPerColumn(Draft, 9, "Yes"))
    .Range("E4").Value = Str(CountPerColumn(Draft, 4, "Female", 9, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 4, "Female", 9, "Yes"), CountPerColumn(Draft, 9, "Yes"))
    End With

    'Age
    Draft.Activate

    RemoveFilter

    DSMBS.Range("B6").Value = Str(Round(WorksheetFunction.Average(Draft.Range("G2:G" & dict("Draft_LR"))), 2)) & " (" & Trim(Str(Round(WorksheetFunction.StDev(Draft.Range("G2:G" & dict("Draft_LR"))), 2))) & ")"
    DSMBS.Range("B7").Value = WorksheetFunction.Median(Draft.Range("G2:G" & dict("Draft_LR")))
    DSMBS.Range("B8").Value = Str(WorksheetFunction.Min(Draft.Range("G2:G" & dict("Draft_LR")))) & " -" & Str(WorksheetFunction.Max(Draft.Range("G2:G" & dict("Draft_LR"))))
    
    Draft.Range("A1").AutoFilter Field:=10, Criteria1:="No"
    DSMBS.Range("C6").Value = Str(Round(WorksheetFunction.Average(Draft.Range("G2:G" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)))) & " (" & Trim(Str(Round(WorksheetFunction.StDev(Draft.Range("G2:G" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)), 2))) & ")"
    DSMBS.Range("C7").Value = WorksheetFunction.Median(Draft.Range("G2:G" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible))
    DSMBS.Range("C8").Value = Str(WorksheetFunction.Min(Draft.Range("G2:G" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible))) & " -" & Str(WorksheetFunction.Max(Draft.Range("G2:G" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)))
    
    RemoveFilter
    Draft.Range("A1").AutoFilter Field:=10, Criteria1:="Yes"
    DSMBS.Range("D6").Value = Str(Round(WorksheetFunction.Average(Draft.Range("G2:G" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)))) & " (" & Trim(Str(Round(WorksheetFunction.StDev(Draft.Range("G2:G" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)), 2))) & ")"
    DSMBS.Range("D7").Value = WorksheetFunction.Median(Draft.Range("G2:G" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible))
    DSMBS.Range("D8").Value = Str(WorksheetFunction.Min(Draft.Range("G2:G" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible))) & " -" & Str(WorksheetFunction.Max(Draft.Range("G2:G" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)))
    
    RemoveFilter
    Draft.Range("A1").AutoFilter Field:=9, Criteria1:="Yes"
    DSMBS.Range("E6").Value = Str(Round(WorksheetFunction.Average(Draft.Range("G2:G" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)))) & " (" & Trim(Str(Round(WorksheetFunction.StDev(Draft.Range("G2:G" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)), 2))) & ")"
    DSMBS.Range("E7").Value = WorksheetFunction.Median(Draft.Range("G2:G" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible))
    DSMBS.Range("E8").Value = Str(WorksheetFunction.Min(Draft.Range("G2:G" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible))) & " -" & Str(WorksheetFunction.Max(Draft.Range("G2:G" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)))
    RemoveFilter
    
    On Error GoTo 0
    'Race
    With DSMBS
    For i = 10 To 18
    
    .Range("B" & i).Value = Str(CountPerColumn(Draft, 8, .Range("A" & i).Value)) & " " & CountToPercent(CountPerColumn(Draft, 8, .Range("A" & i).Value), dict("DM_LR") - 1)

    .Range("C" & i).Value = Str(CountPerColumn(Draft, 8, .Range("A" & i).Value, 10, "No")) & " " & CountToPercent(CountPerColumn(Draft, 8, .Range("A" & i).Value, 10, "No"), CountPerColumn(Draft, 10, "No"))
    
    .Range("D" & i).Value = Str(CountPerColumn(Draft, 8, .Range("A" & i).Value, 10, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 8, .Range("A" & i).Value, 10, "Yes"), CountPerColumn(Draft, 10, "Yes"))
    
    .Range("E" & i).Value = Str(CountPerColumn(Draft, 8, .Range("A" & i).Value, 9, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 8, .Range("A" & i).Value, 9, "Yes"), CountPerColumn(Draft, 9, "Yes"))

    Next i
    
    End With
    
    dict("RowNum") = 12
    DSMBS.Activate
    DSMBS.Range("A1").Select
    Call FormatTable2
    
    Draft.UsedRange.Columns.AutoFit

    Call T2(DSMBWS, DMS, IES, INFS, DLAS, DIAGS, AES, RSS, ITLFS, DSMBS, Draft, dict)
    
End Sub
Sub T2(DSMBWS, DMS, IES, INFS, DLAS, DIAGS, AES, RSS, ITLFS, DSMBS, Draft, dict)
    Draft.Activate
    RemoveFilter
    DSMBS.Activate
    Draft.Range("A1:J" & dict("Draft_LR")).Copy DSMBS.Range("G1")
    DSMBS.Range("G1").Select
    Call FormatTable
    DSMBS.Range("G1").Select
    Call FormatTable2
    
    dict("LC") = L2N("P")
    
     'Remove All Draft Sheet
    Application.DisplayAlerts = False
    Sheets("Draft").Delete
    Application.DisplayAlerts = True

    'Add clean Draft sheet
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(Sheets.Count)
    Set Draft = ActiveWorkbook.Sheets(Sheets.Count)
    ActiveWorkbook.Sheets(Sheets.Count).Name = "Draft"
    
    
    Call T3(DSMBWS, DMS, IES, INFS, DLAS, DIAGS, AES, RSS, ITLFS, DSMBS, Draft, dict)
End Sub

Sub T3(DSMBWS, DMS, IES, INFS, DLAS, DIAGS, AES, RSS, ITLF, DSMBS, Draft, dict)
Dim LC As Long

'Add dict values
LC = dict("LC")

'1)Preparing data in Draft Sheet
'Header
INFS.Activate
RemoveFilter
INFS.Range("A1:AV" & dict("INF_LR")).Sort Key1:=Range("D1"), _
                                          Order1:=xlAscending, Header:=xlYes, _
                                          Key2:=Range("E1"), _
                                          Order1:=xlAscending, Header:=xlYes





With Draft
    .Range("L1").Value = "Subject ID"
    .Range("M1").Value = "Study Day"
    .Range("N1").Value = "Date of Infusion"
    .Range("O1").Value = "Total CART19-IL18 Cell Dose"
    .Range("P1").Value = "Total CART19-IL18 Cell Dose"
    .Range("Q1").Value = "Total Cell Dose Administered"
    .Range("R1").Value = "Total Cell Dose Administered"
    .Range("S1").Value = "%scFv Flow"
    .Range("T1").Value = "Met Target %scFv (Y/N)"
End With

    INFS.Range("D2:E" & dict("INF_LR")).Copy Draft.Range("L2")
    INFS.Range("AE2:AI" & dict("INF_LR")).Copy Draft.Range("N2")
    INFS.Range("AK2:AK" & dict("INF_LR")).Copy Draft.Range("S2")

DSMBS.Activate
Application.DisplayAlerts = False
With DSMBS
    .Cells(1, LC + 2).Value = "Subject ID"
    .Cells(2, LC + 2).Value = "Subject ID"
    .Cells(1, LC + 3).Value = "Study Day"
    .Cells(2, LC + 3).Value = "Study Day"
    .Cells(1, LC + 4).Value = "Date of Infusion"
    .Cells(2, LC + 4).Value = "Date of Infusion"
    .Cells(1, LC + 5).Value = "Cells Infused"
    .Cells(1, LC + 6).Value = "Cells Infused"
    .Cells(1, LC + 7).Value = "Cells Infused"
    .Cells(1, LC + 8).Value = "Cells Infused"
    .Cells(2, LC + 5).Value = "Target Cell Dose*"
    .Cells(2, LC + 6).Value = "Total CART19-IL18 Cell Dose"
    .Cells(2, LC + 7).Value = "Total Cell Dose Administered"
    .Cells(2, LC + 8).Value = "Met Target"
    .Cells(1, LC + 9).Value = "Transduction Efficiency"
    .Cells(1, LC + 10).Value = "Transduction Efficiency"
    .Cells(2, LC + 9).Value = "%scFv Flow"
    .Cells(2, LC + 10).Value = "Met Target %scFv (Y/N)"


End With


On Error Resume Next

For i = 2 To dict("INF_LR")
    DSMBS.Cells(i + 1, LC + 2).Value = Draft.Range("L" & i).Value
    DSMBS.Cells(i + 1, LC + 3).Value = Draft.Range("M" & i).Value
    DSMBS.Cells(i + 1, LC + 4).Value = Draft.Range("N" & i).Value
    DSMBS.Cells(i + 1, LC + 5).Value = "Manually entered"
    DSMBS.Cells(i + 1, LC + 5).Interior.Color = vbYellow
    DSMBS.Cells(i + 1, LC + 6).Value = Draft.Range("O" & i).Value & "x10^" & Draft.Range("P" & i).Value
    DSMBS.Cells(i + 1, LC + 7).Value = Draft.Range("Q" & i).Value & "x10^" & Draft.Range("R" & i).Value
    DSMBS.Cells(i + 1, LC + 8).Value = "Manually entered"
    DSMBS.Cells(i + 1, LC + 8).Interior.Color = vbYellow
    DSMBS.Cells(i + 1, LC + 9).Value = Draft.Range("S" & i).Value & "%"
    If Draft.Range("S" & i).Value > 2 Then
        DSMBS.Cells(i + 1, LC + 10).Value = "Y"
        Else
        DSMBS.Cells(i + 1, LC + 10).Value = "N"
    End If
    'DSMBS.Cells(i, LC + 11).Interior.Color = vbYellow
Next i

On Error GoTo 0

With DSMBS
    .Range(N2L(LC + 2) & 1).Select
    Call FormatTable
    .Range(N2L(LC + 2) & 1).Select
    Call FormatTable2
    .Range(N2L(LC + 2) & "1:" & N2L(LC + 2) & "2").Merge
    .Range(N2L(LC + 2) & "1:" & N2L(LC + 2) & "2").VerticalAlignment = xlCenter
    .Range(N2L(LC + 3) & "1:" & N2L(LC + 3) & "2").Merge
    .Range(N2L(LC + 3) & "1:" & N2L(LC + 3) & "2").VerticalAlignment = xlCenter
    .Range(N2L(LC + 4) & "1:" & N2L(LC + 4) & "2").Merge
    .Range(N2L(LC + 4) & "1:" & N2L(LC + 4) & "2").VerticalAlignment = xlCenter
    .Range(N2L(LC + 5) & "1:" & N2L(LC + 8) & "1").Merge
    .Range(N2L(LC + 5) & "1:" & N2L(LC + 8) & "1").HorizontalAlignment = xlCenter
    .Range(N2L(LC + 9) & "1:" & N2L(LC + 10) & "1").Merge
    .Range(N2L(LC + 9) & "1:" & N2L(LC + 10) & "1").HorizontalAlignment = xlCenter
End With

Application.DisplayAlerts = True

dict("LC") = L2N("Z")

'Remove All Draft Sheet
Application.DisplayAlerts = False
Sheets("Draft").Delete
Application.DisplayAlerts = True

'Add clean Draft sheet
ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(Sheets.Count)
Set Draft = ActiveWorkbook.Sheets(Sheets.Count)
ActiveWorkbook.Sheets(Sheets.Count).Name = "Draft"

Call T4_4(DSMBWS, DMS, IES, INFS, DLAS, DIAGS, AES, RSS, ITLF, DSMBS, Draft, dict)

'Call T4(DMS, IES, INFS, DLAS, DIAGS, AES, RS, DSMBS, Draft, dict)


End Sub
Sub T4_4(DSMBWS, DMS, IES, INFS, DLAS, DIAGS, AES, RSS, ITLF, DSMBS, Draft, dict)

Application.DisplayAlerts = False
Dim LC As Long
Dim FC As Long
Dim pos, arr, val
Dim TPArray() As Variant

TPArray = Array("Day 28", "Month 2", "Month 3", "Month 4", "Month 5", "Month 6", "Month 9", "Month 12", "Month 18", "Month 24", "Month 30", "Month 36", "Month 42", "Month 48", "Month 54", "Month 60", "Year 6", "Year 7", "Year 8", "Year 9", "Year 10", "Year 11", "Year 12", "Year 13", "Year 14", "Year 15", _
                "Day 28-R", "Month 2-R", "Month 3-R", "Month 4-R", "Month 5-R", "Month 6-R", "Month 9-R", "Month 12-R", "Month 18-R", "Month 24-R", "Month 30-R", "Month 36-R", "Month 42-R", "Month 48-R", "Month 54-R", "Month 60-R", "Year 6-R", "Year 7-R", "Year 8-R", "Year 9-R", "Year 10-R", "Year 11-R", "Year 12-R", "Year 13-R", "Year 14-R", "Year 15-R", "Unscheduled")



'Add dict values
FC = dict("LC") + 2

'AE unique count
AES.Activate
RemoveFilter
AES.Range("A1").AutoFilter Field:=39, Criteria1:="AE"
Draft.Range("A1").Value = "AE Unique"
Call CopyUnique(AES, "D2:D" & dict("AE_LR"), Draft, "A2")

'SAE unique count
AES.Activate
RemoveFilter
AES.Range("A1").AutoFilter Field:=39, Criteria1:="SAE"
Draft.Range("B1").Value = "SAE Unique"
Call CopyUnique(AES, "D2:D" & dict("AE_LR"), Draft, "B2")



'-------->>>>  PART1 HEADER:
DSMBS.Cells(1, FC).Value = "Adverse Events (N=" & CountPerColumn(Draft, 1, "*") & ")"
DSMBS.Cells(1, FC + 1).Value = "Adverse Events (N=" & CountPerColumn(Draft, 1, "*") & ")"
DSMBS.Cells(1, FC + 2).Value = "Serious Adverse Events (N=" & CountPerColumn(Draft, 1, "*") & ")"
DSMBS.Cells(1, FC + 3).Value = "Serious Adverse Events (N=" & CountPerColumn(Draft, 1, "*") & ")"
DSMBS.Cells(2, FC).Value = "Yes"
DSMBS.Cells(3, FC).Value = "No"
DSMBS.Cells(2, FC + 2).Value = "Yes"
DSMBS.Cells(3, FC + 2).Value = "No"



'PART1 VALUE:
DSMBS.Cells(2, FC + 1).Value = CountPerColumn(Draft, 1, "*") & CountToPercent(CountPerColumn(Draft, 1, "*"), CountPerColumn(Draft, 1, "*"))
DSMBS.Cells(3, FC + 1).Value = 0 & CountToPercent(0, CountPerColumn(Draft, 1, "*"))
DSMBS.Cells(2, FC + 3).Value = CountPerColumn(Draft, 2, "*") & CountToPercent(CountPerColumn(Draft, 2, "*"), CountPerColumn(Draft, 1, "*"))
DSMBS.Cells(3, FC + 3).Value = (CountPerColumn(Draft, 1, "*") - CountPerColumn(Draft, 2, "*")) & CountToPercent(CountPerColumn(Draft, 1, "*") - CountPerColumn(Draft, 2, "*"), CountPerColumn(Draft, 1, "*"))


'PART 2 Draft
RSS.Activate
RemoveFilter

Draft.Range("C1").Value = "NHL ID"
Draft.Range("D1").Value = "Time Point"
Draft.Range("E1").Value = "Phase"
Draft.Range("F1").Value = "PET response"
Draft.Range("G1").Value = "PET Score"
Draft.Range("H1").Value = "CT response"
Draft.Range("I1").Value = "CT Score"
Draft.Range("J1").Value = "NHL Unique ID"
Draft.Range("K1").Value = "Best PET Score"
Draft.Range("L1").Value = "Best CT Score"
Draft.Range("W1").Value = "Timepoint Score"
RSS.Range("A1").AutoFilter Field:=28, Criteria1:="Non-Hodgkin Lymphoma"
Call CopyUnique(RSS, "D2:D" & dict("RS_LR"), Draft, "J2")


RemoveFilter
RSS.Range("A1").AutoFilter Field:=28, Criteria1:="Non-Hodgkin Lymphoma"
RSS.Range("D2:D" & dict("RS_LR")).Copy Draft.Range("C2")
For i = 2 To dict("RS_LR")
    If IsEmpty(RSS.Range("Y" & i)) Then
        Draft.Range("D" & i).Value = RSS.Range("AD" & i).Value
    Else
        Draft.Range("D" & i).Value = RSS.Range("Y" & i).Value
    End If
Next i

RSS.Range("AC2:AC" & dict("RS_LR")).Copy Draft.Range("E2")
RSS.Range("AJ2:AJ" & dict("RS_LR")).Copy Draft.Range("F2")
RSS.Range("AL2:AL" & dict("RS_LR")).Copy Draft.Range("H2")


'SCORE CALCULATION

For i = 2 To dict("RS_LR")
'PET SCORE
    If Draft.Cells(i, L2N("F")).Value = "Complete Metabolic Response (CMR)" Then
        Draft.Cells(i, L2N("G")).Value = 5
    ElseIf Draft.Cells(i, L2N("F")).Value = "Partial Metabolic Response (PMR)" Then
        Draft.Cells(i, L2N("G")).Value = 4
    ElseIf Draft.Cells(i, L2N("F")).Value = "No Metabolic Response (NMR)" Then
        Draft.Cells(i, L2N("G")).Value = 3
    ElseIf Draft.Cells(i, L2N("F")).Value = "Indeterminate Response (IR)" Then
        Draft.Cells(i, L2N("G")).Value = 2
    ElseIf Draft.Cells(i, L2N("F")).Value = "Progressive Metabolic Disease (PMD)" Then
        Draft.Cells(i, L2N("G")).Value = 1
    Else
        Draft.Cells(i, L2N("G")).Value = 0
    End If
'CT SCORE
    If Draft.Cells(i, L2N("H")).Value = "Complete Radiologic Response (CR)" Then
        Draft.Cells(i, L2N("I")).Value = 5
    ElseIf Draft.Cells(i, L2N("H")).Value = "Partial Response (PR)" Then
        Draft.Cells(i, L2N("I")).Value = 4
    ElseIf Draft.Cells(i, L2N("H")).Value = "Stable Disease (SD)" Then
        Draft.Cells(i, L2N("I")).Value = 3
    ElseIf Draft.Cells(i, L2N("H")).Value = "Indeterminate Response (IR)" Then
        Draft.Cells(i, L2N("I")).Value = 2
    ElseIf Draft.Cells(i, L2N("H")).Value = "Progressive Disease (PD)" Then
        Draft.Cells(i, L2N("I")).Value = 1
    Else
        Draft.Cells(i, L2N("I")).Value = 0
    End If
    
    pos = Application.Match(Draft.Cells(i, L2N("D")).Value, TPArray, False)
    If Not IsError(pos) Then
        Draft.Cells(i, L2N("W")).Value = pos
    End If
    
    
Next i




For i = 2 To CountPerColumn(Draft, L2N("J"), "*") + 1
    RemoveFilter
    Draft.Range("A1").AutoFilter Field:=3, Criteria1:=Draft.Cells(i, L2N("J")).Value
    Draft.Range("A1").AutoFilter Field:=5, Criteria1:="Primary Treatment"
    Draft.Cells(i, L2N("K")).Value = Application.Max(Draft.Range("G:G").SpecialCells(xlCellTypeVisible))
    Draft.Cells(i, L2N("L")).Value = Application.Max(Draft.Range("I:I").SpecialCells(xlCellTypeVisible))
Next i




'-------->>>>>   PART2 HEADER:
DSMBS.Cells(4, FC).Value = "NHL Subject Response (N=" & CountPerColumn(Draft, L2N("J"), "*") & ")"
DSMBS.Cells(5, FC).Value = "PET-Based Response"
DSMBS.Cells(5, FC + 2).Value = "CT-Based Response"

'BEST OVERAL RESPONSE
DSMBS.Cells(6, FC).Value = "Best Overall Response (BOR)"
'PET-based NHL
DSMBS.Cells(7, FC).Value = "Complete Metabolic Response (CMR)"
DSMBS.Cells(8, FC).Value = "Partial Metabolic Response (PMR)"
DSMBS.Cells(9, FC).Value = "No Metabolic Response (NMR)"
DSMBS.Cells(10, FC).Value = "Indeterminate Response (IR)"
DSMBS.Cells(11, FC).Value = "Progressive Metabolic Disease (PMD)"
DSMBS.Cells(12, FC).Value = "Not Assessed"
'CT-Based NHL
DSMBS.Cells(7, FC + 2).Value = "Complete Radiologic Response (CR)"
DSMBS.Cells(8, FC + 2).Value = "Partial Response (PR)"
DSMBS.Cells(9, FC + 2).Value = "Stable Disease (SD)"
DSMBS.Cells(10, FC + 2).Value = "Indeterminate Response (IR)"
DSMBS.Cells(11, FC + 2).Value = "Progressive Disease (PD)"
DSMBS.Cells(12, FC + 2).Value = "Not Assessed"

'3 MONTH RESPONSE
DSMBS.Cells(13, FC).Value = "Overall Response Rate (ORR) at Month 3"
'PET-based NHL
DSMBS.Cells(14, FC).Value = "Complete Metabolic Response (CMR)"
DSMBS.Cells(15, FC).Value = "Partial Metabolic Response (PMR)"
DSMBS.Cells(16, FC).Value = "No Metabolic Response (NMR)"
DSMBS.Cells(17, FC).Value = "Indeterminate Response (IR)"
DSMBS.Cells(18, FC).Value = "Progressive Metabolic Disease (PMD)"
DSMBS.Cells(19, FC).Value = "Not Assessed"
'CT-Based NHL
DSMBS.Cells(14, FC + 2).Value = "Complete Radiologic Response (CR)"
DSMBS.Cells(15, FC + 2).Value = "Partial Response (PR)"
DSMBS.Cells(16, FC + 2).Value = "Stable Disease (SD)"
DSMBS.Cells(17, FC + 2).Value = "Indeterminate Response (IR)"
DSMBS.Cells(18, FC + 2).Value = "Progressive Disease (PD)"
DSMBS.Cells(19, FC + 2).Value = "Not Assessed"

'---->>>>PART2 VALUES

'BOR
'PET-Based NHL

DSMBS.Cells(7, FC + 1).Value = CountPerColumn(Draft, L2N("K"), "5") & CountToPercent(CountPerColumn(Draft, L2N("K"), "5"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(8, FC + 1).Value = CountPerColumn(Draft, L2N("K"), "4") & CountToPercent(CountPerColumn(Draft, L2N("K"), "4"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(9, FC + 1).Value = CountPerColumn(Draft, L2N("K"), "3") & CountToPercent(CountPerColumn(Draft, L2N("K"), "3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(10, FC + 1).Value = CountPerColumn(Draft, L2N("K"), "2") & CountToPercent(CountPerColumn(Draft, L2N("K"), "2"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(11, FC + 1).Value = CountPerColumn(Draft, L2N("K"), "1") & CountToPercent(CountPerColumn(Draft, L2N("K"), "1"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(12, FC + 1).Value = CountPerColumn(Draft, L2N("K"), "0") & CountToPercent(CountPerColumn(Draft, L2N("K"), "0"), CountPerColumn(Draft, L2N("J"), "*"))

'CT-Based NHL
DSMBS.Cells(7, FC + 3).Value = CountPerColumn(Draft, L2N("L"), "5") & CountToPercent(CountPerColumn(Draft, L2N("L"), "5"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(8, FC + 3).Value = CountPerColumn(Draft, L2N("L"), "4") & CountToPercent(CountPerColumn(Draft, L2N("L"), "4"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(9, FC + 3).Value = CountPerColumn(Draft, L2N("L"), "3") & CountToPercent(CountPerColumn(Draft, L2N("L"), "3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(10, FC + 3).Value = CountPerColumn(Draft, L2N("L"), "2") & CountToPercent(CountPerColumn(Draft, L2N("L"), "2"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(11, FC + 3).Value = CountPerColumn(Draft, L2N("L"), "1") & CountToPercent(CountPerColumn(Draft, L2N("L"), "1"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(12, FC + 3).Value = CountPerColumn(Draft, L2N("L"), "0") & CountToPercent(CountPerColumn(Draft, L2N("L"), "0"), CountPerColumn(Draft, L2N("J"), "*"))

'3 MONTH RESPONSE
'PET-based NHL
DSMBS.Cells(14, FC + 1).Value = CountPerColumn(Draft, L2N("G"), "5", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("G"), "5", L2N("D"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(15, FC + 1).Value = CountPerColumn(Draft, L2N("G"), "4", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("G"), "4", L2N("D"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(16, FC + 1).Value = CountPerColumn(Draft, L2N("G"), "3", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("G"), "3", L2N("D"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(17, FC + 1).Value = CountPerColumn(Draft, L2N("G"), "2", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("G"), "2", L2N("D"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(18, FC + 1).Value = CountPerColumn(Draft, L2N("G"), "1", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("G"), "1", L2N("D"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(19, FC + 1).Value = CountPerColumn(Draft, L2N("G"), "0", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("G"), "0", L2N("D"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))

'CT-Based NHL
DSMBS.Cells(14, FC + 3).Value = CountPerColumn(Draft, L2N("I"), "5", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("I"), "5", L2N("D"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(15, FC + 3).Value = CountPerColumn(Draft, L2N("I"), "4", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("I"), "4", L2N("D"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(16, FC + 3).Value = CountPerColumn(Draft, L2N("I"), "3", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("I"), "3", L2N("D"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(17, FC + 3).Value = CountPerColumn(Draft, L2N("I"), "2", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("I"), "2", L2N("D"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(18, FC + 3).Value = CountPerColumn(Draft, L2N("I"), "1", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("I"), "1", L2N("D"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(19, FC + 3).Value = CountPerColumn(Draft, L2N("I"), "0", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("I"), "0", L2N("D"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))



'==========================================================================================================================================================================================================

'PART3 CLL

'PART 3 Draft
RSS.Activate
RemoveFilter

Draft.Range("M1").Value = "CLL ID"
Draft.Range("N1").Value = "Time Point"
Draft.Range("O1").Value = "Phase"
Draft.Range("P1").Value = "PET response"
Draft.Range("Q1").Value = "PET Score"
Draft.Range("R1").Value = "CT response"
Draft.Range("S1").Value = "CT Score"
Draft.Range("T1").Value = "CLL Unique ID"
Draft.Range("U1").Value = "Best PET Score"
Draft.Range("V1").Value = "Best CT Score"

RSS.Range("A1").AutoFilter Field:=28, Criteria1:="Chronic Lymphocytic Leukemia"
RSS.Range("AB:AB").SpecialCells(xlCellTypeVisible).Copy

With Draft
    .Range("T1").PasteSpecial xlPasteValues
    .Range("T1", .Range("T1").End(xlDown)).RemoveDuplicates 1, xlNo
    .Range("T1", .Range("T1").End(xlDown)).Sort Key1:=.Range("T1"), Order1:=xlAscending, Header:=xlNo
End With

RSS.Activate
RemoveFilter
RSS.Range("A1").AutoFilter Field:=28, Criteria1:="Chronic Lymphocytic Leukemia"
On Error Resume Next
RSS.Range("D2:D" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Copy Draft.Range("M2")
RSS.Range("Y2:Y" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Copy Draft.Range("N2")
RSS.Range("AC2:AC" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Copy Draft.Range("O2")
RSS.Range("AE2:AE" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Copy Draft.Range("P2")
RSS.Range("AG2:AG" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Copy Draft.Range("R2")
On Error GoTo 0

'SCORE CALCULATION

For i = 2 To CountPerColumn(Draft, L2N("T"), "*")
'PET SCORE
    If Draft.Cells(i, L2N("P")).Value = "Complete Remission (CR)" Then
        Draft.Cells(i, L2N("Q")).Value = 5
    ElseIf Draft.Cells(i, L2N("P")).Value = "Complete Remission with Incomplete Marrow Recovery (CRi)" Then
        Draft.Cells(i, L2N("Q")).Value = 4
    ElseIf Draft.Cells(i, L2N("P")).Value = "Partial Remission (PR)" Then
        Draft.Cells(i, L2N("Q")).Value = 3
    ElseIf Draft.Cells(i, L2N("P")).Value = "Stable Disease (SD)" Then
        Draft.Cells(i, L2N("Q")).Value = 2
    ElseIf Draft.Cells(i, L2N("P")).Value = "Progressive Disease (PD)" Then
        Draft.Cells(i, L2N("Q")).Value = 1
    Else
        Draft.Cells(i, L2N("Q")).Value = 0
    End If
'CT SCORE
    If Draft.Cells(i, L2N("R")).Value = "Complete Radiologic Response (CR)" Then
        Draft.Cells(i, L2N("S")).Value = 4
    ElseIf Draft.Cells(i, L2N("R")).Value = "Partial Response (PR)" Then
        Draft.Cells(i, L2N("S")).Value = 3
    ElseIf Draft.Cells(i, L2N("R")).Value = "Stable Disease (SD)" Then
        Draft.Cells(i, L2N("S")).Value = 2
    ElseIf Draft.Cells(i, L2N("R")).Value = "Progressive Disease (PD)" Then
        Draft.Cells(i, L2N("S")).Value = 1
    Else
        Draft.Cells(i, L2N("S")).Value = 0
    End If
Next i




For i = 2 To CountPerColumn(Draft, L2N("T"), "*") + 1
    RemoveFilter
    Draft.Range("A1").AutoFilter Field:=13, Criteria1:=Draft.Cells(i, L2N("T")).Value
    Draft.Cells(i, L2N("U")).Value = Application.Max(Draft.Range("Q:Q").SpecialCells(xlCellTypeVisible))
    Draft.Cells(i, L2N("V")).Value = Application.Max(Draft.Range("S:S").SpecialCells(xlCellTypeVisible))
Next i




'-------->>>>>   PART2 HEADER:
DSMBS.Cells(20, FC).Value = "CLL Subject Response (N=" & CountPerColumn(Draft, L2N("T"), "*") & ")"
DSMBS.Cells(21, FC).Value = "Overall Response"
DSMBS.Cells(21, FC + 2).Value = "Bone Marrow Response"

'BEST OVERAL RESPONSE
DSMBS.Cells(22, FC).Value = "Best Overall Response (BOR)"
'PET-based NHL
DSMBS.Cells(23, FC).Value = "Complete Remission (CR)"
DSMBS.Cells(24, FC).Value = "Complete Remission with Incomplete Marrow Recovery (CRi)"
DSMBS.Cells(25, FC).Value = "Partial Remission (PR)"
DSMBS.Cells(26, FC).Value = "Stable Disease (SD)"
DSMBS.Cells(27, FC).Value = "Progressive Disease (PD)"
DSMBS.Cells(28, FC).Value = "Not Assessed"
'CT-Based NHL
DSMBS.Cells(23, FC + 2).Value = "Complete Radiologic Response (CR)"
DSMBS.Cells(24, FC + 2).Value = "Partial Response (PR)"
DSMBS.Cells(25, FC + 2).Value = "Stable Disease (SD)"
DSMBS.Cells(26, FC + 2).Value = "Progressive Disease (PD)"
DSMBS.Cells(27, FC + 2).Value = "Not Assessed"
DSMBS.Cells(28, FC + 2).Interior.Color = 48
DSMBS.Cells(28, FC + 3).Interior.Color = 48
'3 MONTH RESPONSE
DSMBS.Cells(29, FC).Value = "Overall Response Rate (ORR) at Month 3"
'PET-based NHL
DSMBS.Cells(30, FC).Value = "Complete Remission (CR)"
DSMBS.Cells(31, FC).Value = "Complete Remission with Incomplete Marrow Recovery (CRi)"
DSMBS.Cells(32, FC).Value = "Partial Remission (PR)"
DSMBS.Cells(33, FC).Value = "Stable Disease (SD)"
DSMBS.Cells(34, FC).Value = "Progressive Disease (PD)"
DSMBS.Cells(35, FC).Value = "Not Assessed"
'CT-Based NHL
DSMBS.Cells(30, FC + 2).Value = "Complete Radiologic Response (CR)"
DSMBS.Cells(31, FC + 2).Value = "Partial Response (PR)"
DSMBS.Cells(32, FC + 2).Value = "Stable Disease (SD)"
DSMBS.Cells(33, FC + 2).Value = "Progressive Disease (PD)"
DSMBS.Cells(34, FC + 2).Value = "Not Assessed"
DSMBS.Cells(35, FC + 2).Interior.Color = 48
DSMBS.Cells(35, FC + 3).Interior.Color = 48

'---->>>>PART2 VALUES

'BOR
'Overall Response

DSMBS.Cells(23, FC + 1).Value = CountPerColumn(Draft, L2N("U"), "5") & CountToPercent(CountPerColumn(Draft, L2N("U"), "5"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(24, FC + 1).Value = CountPerColumn(Draft, L2N("U"), "4") & CountToPercent(CountPerColumn(Draft, L2N("U"), "4"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(25, FC + 1).Value = CountPerColumn(Draft, L2N("U"), "3") & CountToPercent(CountPerColumn(Draft, L2N("U"), "3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(26, FC + 1).Value = CountPerColumn(Draft, L2N("U"), "2") & CountToPercent(CountPerColumn(Draft, L2N("U"), "2"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(27, FC + 1).Value = CountPerColumn(Draft, L2N("U"), "1") & CountToPercent(CountPerColumn(Draft, L2N("U"), "1"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(28, FC + 1).Value = CountPerColumn(Draft, L2N("U"), "0") & CountToPercent(CountPerColumn(Draft, L2N("U"), "0"), CountPerColumn(Draft, L2N("J"), "*"))

'Bone marrow Response
DSMBS.Cells(23, FC + 3).Value = CountPerColumn(Draft, L2N("V"), "4") & CountToPercent(CountPerColumn(Draft, L2N("V"), "4"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(24, FC + 3).Value = CountPerColumn(Draft, L2N("V"), "3") & CountToPercent(CountPerColumn(Draft, L2N("V"), "3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(25, FC + 3).Value = CountPerColumn(Draft, L2N("V"), "2") & CountToPercent(CountPerColumn(Draft, L2N("V"), "2"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(26, FC + 3).Value = CountPerColumn(Draft, L2N("V"), "1") & CountToPercent(CountPerColumn(Draft, L2N("V"), "1"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(27, FC + 3).Value = CountPerColumn(Draft, L2N("V"), "0") & CountToPercent(CountPerColumn(Draft, L2N("V"), "0"), CountPerColumn(Draft, L2N("J"), "*"))


'3 MONTH RESPONSE
'Overall Response
DSMBS.Cells(30, FC + 1).Value = CountPerColumn(Draft, L2N("Q"), "5", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("Q"), "5", L2N("N"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(31, FC + 1).Value = CountPerColumn(Draft, L2N("Q"), "4", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("Q"), "4", L2N("N"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(32, FC + 1).Value = CountPerColumn(Draft, L2N("Q"), "3", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("Q"), "3", L2N("N"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(33, FC + 1).Value = CountPerColumn(Draft, L2N("Q"), "2", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("Q"), "2", L2N("N"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(34, FC + 1).Value = CountPerColumn(Draft, L2N("Q"), "1", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("Q"), "1", L2N("N"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(35, FC + 1).Value = CountPerColumn(Draft, L2N("Q"), "0", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("Q"), "0", L2N("N"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))

'Bone marrow Response
DSMBS.Cells(30, FC + 3).Value = CountPerColumn(Draft, L2N("S"), "4", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("S"), "4", L2N("N"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(31, FC + 3).Value = CountPerColumn(Draft, L2N("S"), "3", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("S"), "3", L2N("N"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(32, FC + 3).Value = CountPerColumn(Draft, L2N("S"), "2", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("S"), "2", L2N("N"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(33, FC + 3).Value = CountPerColumn(Draft, L2N("S"), "1", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("S"), "1", L2N("N"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))
DSMBS.Cells(34, FC + 3).Value = CountPerColumn(Draft, L2N("S"), "0", L2N("D"), "Month 3") & CountToPercent(CountPerColumn(Draft, L2N("S"), "0", L2N("N"), "Month 3"), CountPerColumn(Draft, L2N("J"), "*"))




'FORMAT
DSMBS.Activate
DSMBS.Range(N2L(FC) & "1").Select
Call FormatTable
DSMBS.Range(N2L(FC) & "1").Select
Call FormatTable2

DSMBS.Range(N2L(FC) & "1:" & N2L(FC + 1) & "1").Merge
DSMBS.Range(N2L(FC) & "1:" & N2L(FC + 1) & "1").Font.Bold = True
DSMBS.Range(N2L(FC + 2) & "1:" & N2L(FC + 3) & "1").Merge
DSMBS.Range(N2L(FC + 2) & "1:" & N2L(FC + 3) & "1").Font.Bold = True
DSMBS.Range(N2L(FC) & "4:" & N2L(FC + 3) & "4").Merge
DSMBS.Range(N2L(FC) & "4:" & N2L(FC + 3) & "4").Font.Bold = True
DSMBS.Range(N2L(FC) & "5:" & N2L(FC + 1) & "5").Merge
DSMBS.Range(N2L(FC) & "5:" & N2L(FC + 1) & "5").Font.Bold = True
DSMBS.Range(N2L(FC + 2) & "5:" & N2L(FC + 3) & "5").Merge
DSMBS.Range(N2L(FC + 2) & "5:" & N2L(FC + 3) & "5").Font.Bold = True
DSMBS.Range(N2L(FC) & "6:" & N2L(FC + 3) & "6").Merge
DSMBS.Range(N2L(FC) & "6:" & N2L(FC + 3) & "6").Font.Bold = True
DSMBS.Range(N2L(FC) & "13:" & N2L(FC + 3) & "13").Merge
DSMBS.Range(N2L(FC) & "13:" & N2L(FC + 3) & "13").Font.Bold = True
DSMBS.Range(N2L(FC) & "20:" & N2L(FC + 3) & "20").Merge
DSMBS.Range(N2L(FC) & "20:" & N2L(FC + 3) & "20").Font.Bold = True
DSMBS.Range(N2L(FC) & "21:" & N2L(FC + 1) & "21").Merge
DSMBS.Range(N2L(FC) & "21:" & N2L(FC + 1) & "21").Font.Bold = True
DSMBS.Range(N2L(FC + 2) & "21:" & N2L(FC + 3) & "21").Merge
DSMBS.Range(N2L(FC + 2) & "21:" & N2L(FC + 3) & "21").Font.Bold = True
DSMBS.Range(N2L(FC) & "22:" & N2L(FC + 3) & "22").Merge
DSMBS.Range(N2L(FC) & "22:" & N2L(FC + 3) & "22").Font.Bold = True
DSMBS.Range(N2L(FC) & "29:" & N2L(FC + 3) & "29").Merge
DSMBS.Range(N2L(FC) & "29:" & N2L(FC + 3) & "29").Font.Bold = True
Application.DisplayAlerts = True


dict("LC") = FC + 3

Call L4_4(DSMBWS, DMS, IES, INFS, DLAS, DIAGS, AES, RSS, ITLF, DSMBS, Draft, dict)


End Sub
Sub L4_4(DSMBWS, DMS, IES, INFS, DLAS, DIAGS, AES, RSS, ITLF, DSMBS, Draft, dict)

'SETUP BEFORE
Dim LC As Long
Dim FC As Long
Dim Draft2 As Worksheet
Dim myDate As Date
Dim rngFiltered As Range
Dim TPArray() As Variant

TPArray = Array("Day 28", "Month 2", "Month 3", "Month 4", "Month 5", "Month 6", "Month 9", "Month 12", "Month 18", "Month 24", "Month 30", "Month 36", "Month 42", "Month 48", "Month 54", "Month 60", "Year 6", "Year 7", "Year 8", "Year 9", "Year 10", "Year 11", "Year 12", "Year 13", "Year 14", "Year 15", _
                "Day 28-R", "Month 2-R", "Month 3-R", "Month 4-R", "Month 5-R", "Month 6-R", "Month 9-R", "Month 12-R", "Month 18-R", "Month 24-R", "Month 30-R", "Month 36-R", "Month 42-R", "Month 48-R", "Month 54-R", "Month 60-R", "Year 6-R", "Year 7-R", "Year 8-R", "Year 9-R", "Year 10-R", "Year 11-R", "Year 12-R", "Year 13-R", "Year 14-R", "Year 15-R", "Unscheduled")


'Add clean Draft sheet
ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(Sheets.Count)
Set Draft2 = ActiveWorkbook.Sheets(Sheets.Count)
ActiveWorkbook.Sheets(Sheets.Count).Name = "Draft2"
Application.DisplayAlerts = False
Set Draft2 = ActiveWorkbook.Sheets("Draft2")

'Add dict values
FC = dict("LC") + 2



'DRAFT

'DRAFT HEADER
Draft2.Range("A1").Value = "Subject ID"
Draft2.Range("B1").Value = "AE Status"
Draft2.Range("C1").Value = "Translated Status"
Draft2.Range("D1").Value = "Adverse Events (Y/N)"
Draft2.Range("E1").Value = "Serious Adverse Events (Y/N)"
Draft2.Range("F1").Value = "Current PET-Based Response"
Draft2.Range("G1").Value = "Current CT-Based Response"
Draft2.Range("H1").Value = "Study Timepoint"
Draft2.Range("I1").Value = "Best PET-Based Response"
Draft2.Range("J1").Value = "Best CT-Based Response"
Draft2.Range("K1").Value = "PET-Based ORR"
Draft2.Range("L1").Value = "CT-Based ORR"

Draft2.Range("A2:A" & dict("AE_LR")).Value = AES.Range("D2:D" & dict("AE_LR")).Value
Draft2.Range("B2:B" & dict("AE_LR")).Value = AES.Range("AF2:AF" & dict("AE_LR")).Value

For i = 2 To dict("AE_LR")

    
    If Draft2.Range("B" & i).Value = "After Retreatment T cell Administration" Then
        Draft2.Range("C" & i).Value = "Primary Retreatment and Follow-up"
    Else
        Draft2.Range("C" & i).Value = "Primary Treatment and Follow-up"
    End If
    
    
    If Not IsError(Application.Match(Draft2.Range("A" & i).Value, Draft.Range("A2:A" & dict("AE_LR")), 0)) Then
        Draft2.Range("D" & i).Value = "Y"
    Else
        Draft2.Range("D" & i).Value = "N"
    End If
    If Not IsError(Application.Match(Draft2.Range("A" & i).Value, Draft.Range("B2:B" & dict("AE_LR")), 0)) Then
        Draft2.Range("E" & i).Value = "Y"
    Else
        Draft2.Range("E" & i).Value = "N"
    End If

Next i



Draft2.Range("A2:H" & dict("AE_LR")).RemoveDuplicates Array(1, 3), xlNo      'OMG I LOVE THIS SO MUCH

Draft2.Range("A1:L1").End(xlDown).Sort Key1:=Draft2.Range("A1"), Order1:=xlAscending, Header:=xlYes
       
       
       



For i = 2 To CountPerColumn(Draft2, 1, "*") + 1
    
    'CURRENT RESPONSE
    RSS.Activate
    RemoveFilter
    RSS.Range("A1").AutoFilter Field:=4, Criteria1:=Draft2.Range("A" & i).Value
    myDate = Application.Max(RSS.Columns(10).SpecialCells(xlCellTypeVisible))
    RSS.Range("A1").AutoFilter Field:=10, Criteria1:=">=" & myDate, Operator:=xlAnd, Criteria2:="<" & myDate + 1
    
    Set rngFiltered = Nothing
    On Error Resume Next
    Set rngFiltered = RSS.Range("J2:J" & dict("RS_LR")).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Not rngFiltered Is Nothing Then
        If RSS.Range("AJ2:AJ" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Value <> "" Then
            Draft2.Range("F" & i).Value = RSS.Range("AJ2:AJ" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Value
        Else
            Draft2.Range("F" & i).Value = "N/A"
        End If
        If RSS.Range("AL2:AL" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Value <> "" Then
            Draft2.Range("G" & i).Value = RSS.Range("AL2:AL" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Value
        Else
            Draft2.Range("G" & i).Value = "N/A"
        End If
        If RSS.Range("Y2:Y" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Value <> "" Then
            Draft2.Range("H" & i).Value = RSS.Range("Y2:Y" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Value
        Else
            Draft2.Range("H" & i).Value = RSS.Range("AD2:AD" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Value
        End If
    
    Else
        Draft2.Range("F" & i).Value = "Pending"
        Draft2.Range("G" & i).Value = "Pending"
        Draft2.Range("H" & i).Value = "Pending"
    End If

        
    
    
    

    Draft.Activate
    RemoveFilter
    
    
    Draft.Range("A1").AutoFilter Field:=3, Criteria1:=Draft2.Range("A" & i).Value
    If Draft2.Range("C" & i).Value = "Primary Treatment and Follow-up" Then
        Draft.Range("A1").AutoFilter Field:=5, Criteria1:="Primary Treatment"
    ElseIf Draft2.Range("C" & i).Value = "Primary Retreatment and Follow-up" Then
        Draft.Range("A1").AutoFilter Field:=5, Criteria1:="Retreatment"
    End If
    Draft.Range("A1").AutoFilter Field:=7, Criteria1:=Application.Max(Draft.Columns(7).SpecialCells(xlCellTypeVisible))
    Draft.Range("A1").AutoFilter Field:=23, Criteria1:=Application.Min(Draft.Columns(23).SpecialCells(xlCellTypeVisible))
    
    Set rngFiltered = Nothing
    On Error Resume Next
    Set rngFiltered = Draft.Range("G2:G" & dict("RS_LR")).SpecialCells(xlCellTypeVisible)
    If Err > 0 Then
            Draft2.Range("I" & i).Value = "Pending"
            Err.Clear
    Else
        If Draft.Range("G2:G" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Cells.Item(1).Value <> "0" Then
            Draft2.Range("I" & i).Value = Draft.Range("F2:F" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Cells.Item(1).Value & "/" & Draft.Range("D2:D" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Cells.Item(1).Value
        Else
            Draft2.Range("I" & i).Value = "N/A"
        End If
    End If
    'CT
    Draft.Activate
    RemoveFilter
    Draft.Range("A1").AutoFilter Field:=3, Criteria1:=Draft2.Range("A" & i).Value
        If Draft2.Range("C" & i).Value = "Primary Treatment and Follow-up" Then
        Draft.Range("A1").AutoFilter Field:=5, Criteria1:="Primary Treatment"
    ElseIf Draft2.Range("C" & i).Value = "Primary Retreatment and Follow-up" Then
        Draft.Range("A1").AutoFilter Field:=5, Criteria1:="Retreatment"
    End If
    Draft.Range("A1").AutoFilter Field:=9, Criteria1:=Application.Max(Draft.Columns(9).SpecialCells(xlCellTypeVisible))
    Draft.Range("A1").AutoFilter Field:=23, Criteria1:=Application.Min(Draft.Columns(23).SpecialCells(xlCellTypeVisible))
    Set rngFiltered = Nothing
    On Error Resume Next
    Set rngFiltered = Draft.Range("H2:H" & dict("RS_LR")).SpecialCells(xlCellTypeVisible)
    If Err > 0 Then
            Draft2.Range("J" & i).Value = "Pending"
            Err.Clear
    Else
        If Draft.Range("I2:I" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Cells.Item(1).Value <> "0" Then
            Draft2.Range("J" & i).Value = Draft.Range("H2:H" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Cells.Item(1).Value & "/" & Draft.Range("D2:D" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Cells.Item(1).Value
        Else
            Draft2.Range("J" & i).Value = "N/A"
        End If
    End If

    'OVERALL RESPONSE/MONTH 3
    Draft.Activate
    RemoveFilter
    Draft.Range("A1").AutoFilter Field:=3, Criteria1:=Draft2.Range("A" & i).Value
    If Draft2.Range("C" & i).Value = "Primary Treatment and Follow-up" Then
        Draft.Range("A1").AutoFilter Field:=4, Criteria1:="Month 3"
    Else
        Draft.Range("A1").AutoFilter Field:=4, Criteria1:="Month 3-R"
    End If
    Set rngFiltered = Nothing
    On Error Resume Next
    Set rngFiltered = Draft.Range("G2:G" & dict("RS_LR")).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Not rngFiltered Is Nothing Then
        If Draft.Range("F2:F" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Value <> "" Then
            Draft2.Range("K" & i).Value = Draft.Range("F2:F" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Value
        Else
            Draft2.Range("K" & i).Value = "N/A"
        End If
    Else
        Draft2.Range("K" & i).Value = "Pending"
    End If
    'CT
    Draft.Activate
    RemoveFilter
    Draft.Range("A1").AutoFilter Field:=3, Criteria1:=Draft2.Range("A" & i).Value
    If Draft2.Range("C" & i).Value = "Primary Treatment and Follow-up" Then
        Draft.Range("A1").AutoFilter Field:=4, Criteria1:="Month 3"
    Else
        Draft.Range("A1").AutoFilter Field:=4, Criteria1:="Month 3-R"
    End If
    Set rngFiltered = Nothing
    On Error Resume Next
    Set rngFiltered = Draft.Range("I2:I" & dict("RS_LR")).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Not rngFiltered Is Nothing Then
        If Draft.Range("H2:H" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Value <> "" Then
            Draft2.Range("L" & i).Value = Draft.Range("H2:H" & dict("RS_LR")).SpecialCells(xlCellTypeVisible).Cells.Item(1).Value
        Else
            Draft2.Range("L" & i).Value = "N/A"
        End If
    Else
        Draft2.Range("L" & i).Value = "Pending"
    End If
    'LTFU
    If Not IsError(Application.Match(Draft2.Range("A" & i).Value, ITLF.Range("D2:D" & dict("ITLF_LR")), 0)) Then
        Draft2.Range("C" & i).Value = "Long Term Follow-up"
    End If
    'TRANSITIONED TO RETREATMENT
    If Draft2.Range("C" & i).Value = "Primary Retreatment and Follow-up" Then
        Draft2.Range("F" & i - 1 & ":H" & i - 1).Merge
        Draft2.Range("F" & i - 1 & ":H" & i - 1).Value = "Transitioned to Retreatment"
        Draft2.Range("A" & i).Value = Draft2.Range("A" & i).Value & "Retx"
    End If



Next i



'HEADER FOR DSMB LISTING 4.4
DSMBS.Activate
With DSMBS
    .Cells(1, FC).Value = "Subject ID"
    .Cells(1, FC + 1).Value = "Current Response"
    .Cells(1, FC + 2).Value = "Current Response"
    .Cells(1, FC + 3).Value = "Current Response"
    .Cells(1, FC + 4).Value = "Best Response/Timepoint"
    .Cells(1, FC + 5).Value = "Best Response/Timepoint"
    .Cells(1, FC + 6).Value = "Overall Response/Month 3"
    .Cells(1, FC + 7).Value = "Overall Response/Month 3"
    .Cells(1, FC + 8).Value = "Adverse Events (Y/N)"
    .Cells(1, FC + 9).Value = "Serious Adverse Events (Y/N)"
    .Cells(1, FC + 10).Value = "Study Status"

End With

Draft2.Activate
Draft2.Range("A1").Select
FormatTable
Draft2.Range("A1").Select
FormatTable2

Draft2.Range("A1:A" & CountPerColumn(Draft2, 1, "*") + 1).Copy DSMBS.Range(N2L(FC) & 2)
Draft2.Range("F1:L" & CountPerColumn(Draft2, 1, "*") + 1).Copy DSMBS.Range(N2L(FC + 1) & 2)
Draft2.Range("D1:E" & CountPerColumn(Draft2, 1, "*") + 1).Copy DSMBS.Range(N2L(FC + 8) & 2)
Draft2.Range("C1:C" & CountPerColumn(Draft2, 1, "*") + 1).Copy DSMBS.Range(N2L(FC + 10) & 2)



DSMBS.Activate
DSMBS.Range(N2L(FC) & 1).Select
FormatTable
DSMBS.Range(N2L(FC) & 1).Select
FormatTable2

DSMBS.Range(N2L(FC) & "1:" & N2L(FC) & "2").Merge
DSMBS.Range(N2L(FC) & "1:" & N2L(FC) & "2").Font.Bold = True
DSMBS.Range(N2L(FC + 1) & "1:" & N2L(FC + 3) & "1").Merge
DSMBS.Range(N2L(FC + 1) & "1:" & N2L(FC + 3) & "1").Font.Bold = True
DSMBS.Range(N2L(FC + 4) & "1:" & N2L(FC + 5) & "1").Merge
DSMBS.Range(N2L(FC + 4) & "1:" & N2L(FC + 5) & "1").Font.Bold = True
DSMBS.Range(N2L(FC + 6) & "1:" & N2L(FC + 7) & "1").Merge
DSMBS.Range(N2L(FC + 6) & "1:" & N2L(FC + 7) & "1").Font.Bold = True
DSMBS.Range(N2L(FC + 8) & "1:" & N2L(FC + 8) & "2").Merge
DSMBS.Range(N2L(FC + 8) & "1:" & N2L(FC + 8) & "2").Font.Bold = True
DSMBS.Range(N2L(FC + 9) & "1:" & N2L(FC + 9) & "2").Merge
DSMBS.Range(N2L(FC + 9) & "1:" & N2L(FC + 9) & "2").Font.Bold = True
DSMBS.Range(N2L(FC + 10) & "1:" & N2L(FC + 10) & "2").Merge
DSMBS.Range(N2L(FC + 10) & "1:" & N2L(FC + 10) & "2").Font.Bold = True



'Remove All Draft Sheet
Application.DisplayAlerts = False
Sheets("Draft").Delete
Sheets("Draft2").Delete
Application.DisplayAlerts = True


End Sub

Sub T4(DSMBWS, DMS, IES, INFS, DLAS, DIAGS, AES, RSS, ITLF, DSMBS, Draft, dict)
'SAE tables
AES.Activate
RemoveFilter
AES.Range("A1:CU" & dict("AE_LR")).Sort Key1:=Range("D1"), _
                                          Order1:=xlAscending, Header:=xlYes, _
                                          Key2:=Range("AC1"), _
                                          Order1:=xlAscending, Header:=xlYes

With DSMBS
    .Range("Z1").Value = "Subject ID"
    .Range("AA1").Value = "T-Cell Atribution"
    .Range("AB1").Value = "CTCAE Category"
    .Range("AC1").Value = "Toxicity"
    .Range("AD1").Value = "Grade"
    .Range("AE1").Value = "Start Date"
    .Range("AF1").Value = "Stop Date"
    '.Range("APG1").Value = "Duration (days)"
End With
    
    AES.Activate
    AES.Range("A1").AutoFilter Field:=39, Criteria1:="SAE"
    
AES.Range("D2:D" & dict("AE_LR")).Select
Selection.SpecialCells(xlCellTypeVisible).Copy DSMBS.Range("Z2")

AES.Range(dict("AE_Attribution") & "2:" & dict("AE_Attribution") & dict("AE_LR")).Select
Selection.SpecialCells(xlCellTypeVisible).Copy DSMBS.Range("AA2")

AES.Range(dict("AE_CTCAE") & "2:" & dict("AE_CTCAE") & dict("AE_LR")).Select
Selection.SpecialCells(xlCellTypeVisible).Copy DSMBS.Range("AB2")

AES.Range(dict("AE_Toxicity") & "2:" & dict("AE_Toxicity") & dict("AE_LR")).Select
Selection.SpecialCells(xlCellTypeVisible).Copy DSMBS.Range("AC2")

AES.Range(dict("AE_Grade") & "2:" & dict("AE_Grade") & dict("AE_LR")).Select
Selection.SpecialCells(xlCellTypeVisible).Copy DSMBS.Range("AD2")

AES.Range(dict("AE_StartDate") & "2:" & dict("AE_StartDate") & dict("AE_LR")).Select
Selection.SpecialCells(xlCellTypeVisible).Copy DSMBS.Range("AE2")

AES.Range(dict("AE_StopDate") & "2:" & dict("AE_StopDate") & dict("AE_LR")).Select
Selection.SpecialCells(xlCellTypeVisible).Copy DSMBS.Range("AF2")

DSMBS.Activate
DSMBS.Range("Z1").Select
Call FormatTable



'Call T5(DSMBWS, DMS, IES, INFS, DLAS, DIAGS, AES, RSS, ITLF, DSMBS, Draft, dict)


End Sub

Sub T5(DSMBWS, DMS, IES, INFS, DLAS, DIAGS, AES, RSS, ITLF, DSMBS, Draft, dict)
'AE table
AES.Activate
RemoveFilter
AES.Range("A1:CU" & dict("AE_LR")).Sort Key1:=Range("D1"), _
                                          Order1:=xlAscending, Header:=xlYes, _
                                          Key2:=Range("AC1"), _
                                          Order1:=xlAscending, Header:=xlYes

With DSMBS
    .Range("AH1").Value = "Subject ID"
    .Range("AI1").Value = "T-Cell Atribution"
    .Range("AJ1").Value = "CTCAE Category"
    .Range("AK1").Value = "Toxicity"
    .Range("AL1").Value = "Grade"
    .Range("AM1").Value = "Start Date"
    .Range("AN1").Value = "Stop Date"
    '.Range("APG1").Value = "Duration (days)"
End With
    
    AES.Activate
    AES.Range("A1").AutoFilter Field:=39, Criteria1:="AE"
    
AES.Range("D2:D" & dict("AE_LR")).Select
Selection.SpecialCells(xlCellTypeVisible).Copy DSMBS.Range("AH2")

AES.Range(dict("AE_Attribution") & "2:" & dict("AE_Attribution") & dict("AE_LR")).Select
Selection.SpecialCells(xlCellTypeVisible).Copy DSMBS.Range("AI2")

AES.Range(dict("AE_CTCAE") & "2:" & dict("AE_CTCAE") & dict("AE_LR")).Select
Selection.SpecialCells(xlCellTypeVisible).Copy DSMBS.Range("AJ2")

AES.Range(dict("AE_Toxicity") & "2:" & dict("AE_Toxicity") & dict("AE_LR")).Select
Selection.SpecialCells(xlCellTypeVisible).Copy DSMBS.Range("AK2")

AES.Range(dict("AE_Grade") & "2:" & dict("AE_Grade") & dict("AE_LR")).Select
Selection.SpecialCells(xlCellTypeVisible).Copy DSMBS.Range("AL2")

AES.Range(dict("AE_StartDate") & "2:" & dict("AE_StartDate") & dict("AE_LR")).Select
Selection.SpecialCells(xlCellTypeVisible).Copy DSMBS.Range("AM2")

AES.Range(dict("AE_StopDate") & "2:" & dict("AE_StopDate") & dict("AE_LR")).Select
Selection.SpecialCells(xlCellTypeVisible).Copy DSMBS.Range("AN2")

DSMBS.Activate
DSMBS.Range("AH1").Select
Call FormatTable

'Call T6(DMS, IES, INFS, DLAS, DIAGS, AES, RS, DSMBS, Draft, dict)


End Sub
Sub MingComprehensiveFormSummary()

Dim FormWS As Worksheet, FormWB As Workbook, FormLR As Long
Dim EventWS As Worksheet, EventWB As Workbook, EventLR As Long
Dim CurrentWS As Worksheet, CurrentWB As Workbook, CurrentLR As Long
Dim TransWS As Worksheet, TransLR As Long
Dim FormCB As Long
Dim EventCB As Long


Application.ScreenUpdating = False
Application.DisplayAlerts = False

Set CurrentWS = ActiveSheet
Set CurrentWB = ActiveWorkbook
Set FormWB = Workbooks.Open("A:\Testing Environment\Hoang's testing\S15420\form_definition_names_and_descriptions.csv")
Set EventWB = Workbooks.Open("A:\Testing Environment\Hoang's testing\S15420\event_definition_names_and_descriptions.csv")

'copy new sheets over
FormWB.Sheets(1).Copy After:=CurrentWB.Sheets(1)
FormWB.Close SaveChanges:=False
Set FormWS = CurrentWB.Sheets(2)
EventWB.Sheets(1).Copy After:=CurrentWB.Sheets(2)
EventWB.Close SaveChanges:=False
Set EventWS = CurrentWB.Sheets(3)

'Create translation sheet
CurrentWB.Worksheets.Add After:=CurrentWB.Sheets(3)
Set TransWS = CurrentWB.Sheets(4)
CurrentWB.Sheets(4).Name = "Translations"

FormLR = FindLastRowA(FormWS)

EventLR = FindLastRowA(EventWS)

CurrentLR = FindLastRowA(CurrentWS)




FormWS.Activate
FormCB = Application.WorksheetFunction.Max(Range("a:a"))
FormWS.Range("A1").AutoFilter Field:=1, Criteria1:=Str(FormCB)
FormWS.Range("B1:C" & FormLR).Select
Selection.SpecialCells(xlCellTypeVisible).Copy TransWS.Range("A1")

EventWS.Activate
EventCB = Application.WorksheetFunction.Max(Range("a:a"))
EventWS.Range("A1").AutoFilter Field:=1, Criteria1:=Str(EventCB)
EventWS.Range("G1:H" & EventLR).Select
Selection.SpecialCells(xlCellTypeVisible).Copy TransWS.Range("D1")

TransLR = FindLastRowA(TransWS)

CurrentWS.Activate
Range("H1").EntireColumn.Insert
Range("C1").EntireColumn.Insert
Range("I1").Value = "Event Label"
Range("C1").Value = "Form Label"
On Error Resume Next
For i = 2 To CurrentLR
    'Dose Level
    Cells(i, 3).Value = TransWS.Range("B" & WorksheetFunction.Match(Cells(i, 2).Value, TransWS.Range("A1:A" & TransLR), 0))
    Cells(i, 9).Value = TransWS.Range("E" & WorksheetFunction.Match(Cells(i, 8).Value, TransWS.Range("D1:D" & TransLR), 0))
Next i
    
On Error GoTo 0


End Sub

