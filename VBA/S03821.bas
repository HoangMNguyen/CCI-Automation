Attribute VB_Name = "S03821"
Sub DSMB_Report()

Dim DMS As Worksheet, DMLastRow As Long
Dim IES As Worksheet, IELastRow As Long
Dim INFS As Worksheet, INFLastRow As Long
Dim DLAS As Worksheet, DLALastRow As Long
Dim AES As Worksheet, AELastRow As Long
Dim DIAGS As Worksheet, DIAGLastRow As Long
Dim RSS As Worksheet, RSLastRow As Long
Dim DSMBWB As Workbook, DSMBS As Worksheet
Dim Draft As Worksheet
Dim dict As Scripting.Dictionary

Dim RowNum As Long

Set dict = New Scripting.Dictionary

RowNum = 1

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Set DSMBWB = Workbooks.Add
DSMBWB.Activate
MsgBox ("Please select DM WorkSheet")
ImportSheet (1)
Set DMS = DSMBWB.Sheets(1)
MsgBox ("Please select IE WorkSheet")
ImportSheet (2)
Set IES = DSMBWB.Sheets(2)
MsgBox ("Please select INF WorkSheet")
ImportSheet (3)
Set INFS = DSMBWB.Sheets(3)
MsgBox ("Please select AE WorkSheet")
ImportSheet (4)
Set AES = DSMBWB.Sheets(4)
MsgBox ("Please select DLA WorkSheet")
ImportSheet (5)
Set DLAS = DSMBWB.Sheets(5)
MsgBox ("Please select PRDIAG WorkSheet")
ImportSheet (6)
Set DIAGS = DSMBWB.Sheets(6)
MsgBox ("Please select RS WorkSheet")
ImportSheet (7)
Set RSS = DSMBWB.Sheets(7)
'Add DSMB Report sheet
DSMBWB.Sheets("Sheet1").Name = "DSMB Report"
Set DSMBS = DSMBWB.Sheets("DSMB Report")
'Add Draft sheet
DSMBWB.Worksheets.Add After:=DSMBWB.Sheets(8)
Set Draft = DSMBWB.Sheets(9)
DSMBWB.Sheets(9).Name = "Draft"

'Count last row of each sheet and add these values to dict

DMLastRow = FindLastRowA(DMS)
dict.Add "DM_LR", DMLastRow

IELastRow = FindLastRowA(IES)
dict.Add "IE_LR", IELastRow

INFLastRow = FindLastRowA(INFS)
dict.Add "INF_LR", INFLastRow

DLALastRow = FindLastRowA(DLAS)
dict.Add "DLA_LR", DLALastRow

AELastRow = FindLastRowA(AES)
dict.Add "AE_LR", AELastRow

DIAGLastRow = FindLastRowA(DIAGS)
dict.Add "DIAG_LR", DIAGLastRow

RSLastRow = FindLastRowA(RSS)
dict.Add "RS_LR", RSLastRow

dict.Add "RowNum", RowNum
'Adding values for each sheets to dict
dict.Add "DM_LSex", "Z"
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


Call T1(DMS, IES, INFS, DLAS, DIAGS, AES, RS, DSMBS, Draft, dict)



Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub T1(DMS, IES, INFS, DLAS, DIAGS, AES, RS, DSMBS, Draft, dict)

Dim Screenfailed As Long
Dim Enrolled As Long
Dim Infused As Long
Dim MaleTS As Long
Dim BirthDay As Variant
Dim DraftLastRow As Long

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
    Cells(i, 5).Value = DMS.Range(dict("DM_LSex") & WorksheetFunction.Match(Cells(i, 1).Value, DMS.Range("D1:D" & dict("DM_LR")), 0))
    'Sex Assigned at birth
    Cells(i, 6).Value = DMS.Range(dict("DM_LSex") & WorksheetFunction.Match(Cells(i, 1).Value, DMS.Range("D1:D" & dict("DM_LR")), 0))
    'Age
    BirthDay = DMS.Range(dict("DM_Age") & WorksheetFunction.Match(Cells(i, 1).Value, DMS.Range("D1:D" & dict("DM_LR")), 0))
    Cells(i, 7).Value = DateDiff("yyyy", BirthDay, Now)
    'Race (column AC of DM)
    Cells(i, 8).Value = DMS.Range(dict("DM_Race") & WorksheetFunction.Match(Cells(i, 1).Value, DMS.Range("D1:D" & dict("DM_LR")), 0))
    'Infused (column Y of INF)
    Cells(i, 9).Value = INFS.Range(dict("INF_INF") & WorksheetFunction.Match(Cells(i, 1).Value, INFS.Range("D1:D" & dict("INF_LR")), 0))
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
Screenfailed = CountPerColumn(Draft, 8, "No")
DSMBS.Range("C1").Value = "Screen Failed" & vbNewLine & "N=" & Str(Screenfailed)

'count number of Enrolled
Enrolled = CountPerColumn(Draft, 8, "Yes")
DSMBS.Range("D1").Value = "Enrolled" & vbNewLine & "N=" & Str(Enrolled)

'count number of Infused
Infused = CountPerColumn(Draft, 7, "Yes")
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
    With DSMBS
    .Range("B3").Value = Str(CountPerColumn(DMS, 26, "Male")) & " " & CountToPercent(CountPerColumn(DMS, 26, "Male"), dict("DM_LR") - 1)
    .Range("B4").Value = Str(CountPerColumn(DMS, 26, "Female")) & " " & CountToPercent(CountPerColumn(DMS, 26, "Female"), dict("DM_LR") - 1)
    .Range("C3").Value = Str(CountPerColumn(Draft, 4, "Male", 8, "No")) & " " & CountToPercent(CountPerColumn(Draft, 4, "Male", 8, "No"), CountPerColumn(Draft, 8, "No"))
    .Range("C4").Value = Str(CountPerColumn(Draft, 4, "Female", 8, "No")) & " " & CountToPercent(CountPerColumn(Draft, 4, "Female", 8, "No"), CountPerColumn(Draft, 8, "No"))
    .Range("D3").Value = Str(CountPerColumn(Draft, 4, "Male", 8, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 4, "Male", 8, "Yes"), CountPerColumn(Draft, 8, "Yes"))
    .Range("D4").Value = Str(CountPerColumn(Draft, 4, "Female", 8, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 4, "Female", 8, "Yes"), CountPerColumn(Draft, 8, "Yes"))
    .Range("E3").Value = Str(CountPerColumn(Draft, 4, "Male", 7, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 4, "Male", 7, "Yes"), CountPerColumn(Draft, 7, "Yes"))
    .Range("E4").Value = Str(CountPerColumn(Draft, 4, "Female", 7, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 4, "Female", 7, "Yes"), CountPerColumn(Draft, 7, "Yes"))
    End With

    'Age
    Draft.Activate

    RemoveFilter

    DSMBS.Range("B6").Value = Str(Round(WorksheetFunction.Average(Draft.Range("E2:E" & dict("Draft_LR"))), 2)) & " (" & Trim(Str(Round(WorksheetFunction.StDev(Draft.Range("E2:E" & dict("Draft_LR"))), 2))) & ")"
    DSMBS.Range("B7").Value = WorksheetFunction.Median(Draft.Range("E2:E" & dict("Draft_LR")))
    DSMBS.Range("B8").Value = Str(WorksheetFunction.Min(Draft.Range("E2:E" & dict("Draft_LR")))) & " -" & Str(WorksheetFunction.Max(Draft.Range("E2:E" & dict("Draft_LR"))))
    
    Draft.Range("A1").AutoFilter Field:=8, Criteria1:="No"
    DSMBS.Range("C6").Value = Str(Round(WorksheetFunction.Average(Draft.Range("E2:E" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)))) & " (" & Trim(Str(Round(WorksheetFunction.StDev(Draft.Range("E2:E" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)), 2))) & ")"
    DSMBS.Range("C7").Value = WorksheetFunction.Median(Draft.Range("E2:E" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible))
    DSMBS.Range("C8").Value = Str(WorksheetFunction.Min(Draft.Range("E2:E" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible))) & " -" & Str(WorksheetFunction.Max(Draft.Range("E2:E" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)))
    
    RemoveFilter
    Draft.Range("A1").AutoFilter Field:=8, Criteria1:="Yes"
    DSMBS.Range("D6").Value = Str(Round(WorksheetFunction.Average(Draft.Range("E2:E" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)))) & " (" & Trim(Str(Round(WorksheetFunction.StDev(Draft.Range("E2:E" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)), 2))) & ")"
    DSMBS.Range("D7").Value = WorksheetFunction.Median(Draft.Range("E2:E" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible))
    DSMBS.Range("D8").Value = Str(WorksheetFunction.Min(Draft.Range("E2:E" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible))) & " -" & Str(WorksheetFunction.Max(Draft.Range("E2:E" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)))
    
    RemoveFilter
    Draft.Range("A1").AutoFilter Field:=7, Criteria1:="Yes"
    DSMBS.Range("E6").Value = Str(Round(WorksheetFunction.Average(Draft.Range("E2:E" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)))) & " (" & Trim(Str(Round(WorksheetFunction.StDev(Draft.Range("E2:E" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)), 2))) & ")"
    DSMBS.Range("E7").Value = WorksheetFunction.Median(Draft.Range("E2:E" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible))
    DSMBS.Range("E8").Value = Str(WorksheetFunction.Min(Draft.Range("E2:E" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible))) & " -" & Str(WorksheetFunction.Max(Draft.Range("E2:E" & dict("Draft_LR")).SpecialCells(xlCellTypeVisible)))
    RemoveFilter
    
    'Race
    With DSMBS
    .Range("B10").Value = Str(CountPerColumn(Draft, 6, "Caucasian")) & " " & CountToPercent(CountPerColumn(Draft, 6, "Caucasian"), dict("DM_LR") - 1)
    .Range("B11").Value = Str(CountPerColumn(Draft, 6, "African American")) & " " & CountToPercent(CountPerColumn(Draft, 6, "African American"), dict("DM_LR") - 1)
    .Range("B12").Value = Str(CountPerColumn(Draft, 6, "Unknown")) & " " & CountToPercent(CountPerColumn(Draft, 6, "Unknown"), dict("DM_LR") - 1)
    .Range("C10").Value = Str(CountPerColumn(Draft, 6, "Caucasian", 8, "No")) & " " & CountToPercent(CountPerColumn(Draft, 6, "Caucasian", 8, "No"), CountPerColumn(Draft, 8, "No"))
    .Range("C11").Value = Str(CountPerColumn(Draft, 6, "African American", 8, "No")) & " " & CountToPercent(CountPerColumn(Draft, 6, "African American", 8, "No"), CountPerColumn(Draft, 8, "No"))
    .Range("C12").Value = Str(CountPerColumn(Draft, 6, "Unknown", 8, "No")) & " " & CountToPercent(CountPerColumn(Draft, 6, "Unknown", 8, "No"), CountPerColumn(Draft, 8, "No"))
    
    .Range("D10").Value = Str(CountPerColumn(Draft, 6, "Caucasian", 8, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 6, "Caucasian", 8, "Yes"), CountPerColumn(Draft, 8, "Yes"))
    .Range("D11").Value = Str(CountPerColumn(Draft, 6, "African American", 8, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 6, "African American", 8, "Yes"), CountPerColumn(Draft, 8, "Yes"))
    .Range("D12").Value = Str(CountPerColumn(Draft, 6, "Unknown", 8, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 6, "Unknown", 8, "Yes"), CountPerColumn(Draft, 8, "Yes"))
    .Range("E10").Value = Str(CountPerColumn(Draft, 6, "Caucasian", 7, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 6, "Caucasian", 7, "Yes"), CountPerColumn(Draft, 7, "Yes"))
    .Range("E11").Value = Str(CountPerColumn(Draft, 6, "African American", 7, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 6, "African American", 7, "Yes"), CountPerColumn(Draft, 7, "Yes"))
    .Range("E12").Value = Str(CountPerColumn(Draft, 6, "Unknown", 7, "Yes")) & " " & CountToPercent(CountPerColumn(Draft, 6, "Unknown", 7, "Yes"), CountPerColumn(Draft, 7, "Yes"))
    
    
    
    
    
    End With
    
    dict("RowNum") = 12
    DSMBS.Activate
    DSMBS.Range("A1").Select
    Call FormatTable2
    
    Draft.UsedRange.Columns.AutoFit

    Call T2(DMS, IES, INFS, DLAS, DIAGS, AES, RS, DSMBS, Draft, dict)
    
End Sub
Sub T2(DMS, IES, INFS, DLAS, DIAGS, AES, RS, DSMBS, Draft, dict)
    Draft.Activate
    RemoveFilter
    DSMBS.Activate
    Draft.Range("A1:J" & dict("Draft_LR")).Copy DSMBS.Range("G1")
    DSMBS.Range("G1").Select
    Call FormatTable
    DSMBS.Range("G1").Select
    Call FormatTable2
    Call T3(DMS, IES, INFS, DLAS, DIAGS, AES, RS, DSMBS, Draft, dict)
End Sub

Sub T3(DMS, IES, INFS, DLAS, DIAGS, AES, RS, DSMBS, Draft, dict)

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
    .Range("P1").Value = "Subject ID"
    .Range("P2").Value = "Subject ID"
    .Range("Q1").Value = "Study Day"
    .Range("Q2").Value = "Study Day"
    .Range("R1").Value = "Date of Infusion"
    .Range("R2").Value = "Date of Infusion"
    .Range("S1").Value = "Cells Infused"
    .Range("T1").Value = "Cells Infused"
    .Range("U1").Value = "Cells Infused"
    .Range("V1").Value = "Cells Infused"
    .Range("S2").Value = "Target Cell Dose*"
    .Range("T2").Value = "Total CART19-IL18 Cell Dose"
    .Range("U2").Value = "Total Cell Dose Administered"
    .Range("V2").Value = "Met Target"
    .Range("W1").Value = "Transduction Efficiency"
    .Range("X1").Value = "Transduction Efficiency"
    .Range("W2").Value = "%scFv Flow"
    .Range("X2").Value = "Met Target %scFv (Y/N)"
    .Range("P1").Select

End With


On Error Resume Next

For i = 2 To dict("INF_LR")
    DSMBS.Range("P" & i + 1).Value = Draft.Range("L" & i).Value
    DSMBS.Range("Q" & i + 1).Value = Draft.Range("M" & i).Value
    DSMBS.Range("R" & i + 1).Value = Draft.Range("N" & i).Value
    DSMBS.Range("S" & i + 1).Value = "Manually entered"
    DSMBS.Range("S" & i + 1).Interior.Color = vbYellow
    DSMBS.Range("T" & i + 1).Value = Draft.Range("O" & i).Value & "x10^" & Draft.Range("P" & i).Value
    DSMBS.Range("U" & i + 1).Value = Draft.Range("Q" & i).Value & "x10^" & Draft.Range("R" & i).Value
    DSMBS.Range("V" & i + 1).Value = "Manually entered"
    DSMBS.Range("V" & i + 1).Interior.Color = vbYellow
    DSMBS.Range("W" & i + 1).Value = Draft.Range("S" & i).Value & "%"
    DSMBS.Range("X" & i + 1).Value = "Manually entered"
    DSMBS.Range("X" & i + 1).Interior.Color = vbYellow
Next i

On Error GoTo 0

With DSMBS
    .Range("P1").Select
    Call FormatTable
    .Range("P1:P2").Merge
    .Range("P1:P2").VerticalAlignment = xlCenter
    .Range("Q1:Q2").Merge
    .Range("Q1:Q2").VerticalAlignment = xlCenter
    .Range("R1:R2").Merge
    .Range("R1:R2").VerticalAlignment = xlCenter
    .Range("S1:V1").Merge
    .Range("S1:V1").HorizontalAlignment = xlCenter
    .Range("W1:X1").Merge
    .Range("W1:X1").HorizontalAlignment = xlCenter
End With

Application.DisplayAlerts = True

Call T4(DMS, IES, INFS, DLAS, DIAGS, AES, RS, DSMBS, Draft, dict)


End Sub
Sub T4(DMS, IES, INFS, DLAS, DIAGS, AES, RS, DSMBS, Draft, dict)

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



Call T5(DMS, IES, INFS, DLAS, DIAGS, AES, RS, DSMBS, Draft, dict)


End Sub

Sub T5(DMS, IES, INFS, DLAS, DIAGS, AES, RS, DSMBS, Draft, dict)

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

Sub T6(DMS, IES, INFS, DLAS, DIAGS, AES, RS, DSMBS, Draft, dict)



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
Set FormWB = Workbooks.Open("A:\Testing Environment\Hoang's testing\S03821\form_definition_names_and_descriptions.csv")
Set EventWB = Workbooks.Open("A:\Testing Environment\Hoang's testing\S03821\event_definition_names_and_descriptions.csv")

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

On Error Resume Next
For i = 2 To CurrentLR
    'Dose Level
    Cells(i, 3).Value = TransWS.Range("B" & WorksheetFunction.Match(Cells(i, 2).Value, TransWS.Range("A1:A" & TransLR), 0))

Next i
    
On Error GoTo 0

End Sub


