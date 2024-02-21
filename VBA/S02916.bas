Attribute VB_Name = "S02916"

Sub PRIOR_ONC() 'must run the report in linear format, did not remove the blank rows incase there are rows that regimen # is blank but still has data entered.

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
Sub PRIOR_ANP() 'reformat prior ANP v2

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
For i = 16 To 25
    WS1.Cells(8, i).Copy WS2.Cells(1, i - 14)
Next i



'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
'Loop rows
l = 1
For i = 9 To RowNum
    For j = 16 To 115 Step 10
    If WS1.Cells(i, j).Value <> "No" Then
        l = l + 1
        WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
        For k = 0 To 9
       '     If WS2.Cells(i, j).Value <> "No" Then
                WS1.Cells(i, k + j).Copy WS2.Cells(l, k + 2)
        '    End If
        Next k
    End If
    Next j
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat
End Sub

Sub INFUSIONVITALS_COHORT6()

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
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Filtered Infusion Vital Signs"
    Set WS2 = Sheets("Filtered Infusion Vital Signs")
End With
'Set up header for WS2
WS1.Range("C8").Copy WS2.Range("A1")
For i = 19 To 26
    WS1.Cells(8, i).Copy WS2.Cells(1, i - 15) 'ws2 start from column 4 for repeat headers
Next i
'header for other common fields before the repeat items
WS2.Range("B1").Value = "Visit ID"
WS2.Range("C1").Value = "Infusion Number"

'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
'Loop rows
l = 1
For i = 9 To RowNum
    For j = 19 To 178 Step 8
    If WS1.Cells(i, j).Value <> "Not Applicable" Then
        l = l + 1
          'copy subject ID from worksheet 1 (9, 3) to worksheet 2 (2,1) all the way to worksheet 1(last row,3) to worksheet2(last row, 1)
            WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
            WS1.Cells(i, 17).Copy WS2.Cells(l, 2)
            WS1.Cells(i, 18).Copy WS2.Cells(l, 3)
        For k = 0 To 7
                WS1.Cells(i, k + j).Copy WS2.Cells(l, k + 4)
        Next k
    End If
    Next j
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat
End Sub

Sub INFUSIONVITALS_COHORT5()

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
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Filtered Infusion Vital Signs"
    Set WS2 = Sheets("Filtered Infusion Vital Signs")
End With
'Set up header for WS2
WS1.Range("C8").Copy WS2.Range("A1")
For i = 18 To 25
    WS1.Cells(8, i).Copy WS2.Cells(1, i - 16) 'ws2 start from column 2 for repeat headers
Next i

'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
'Loop rows
l = 1
For i = 9 To RowNum
    For j = 18 To 177 Step 8
    If WS1.Cells(i, j).Value <> "Not Applicable" Then
        l = l + 1
          'copy subject ID from worksheet 1 (9, 3) to worksheet 2 (2,1) all the way to worksheet 1(last row,3) to worksheet2(last row, 1)
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

Sub MEDHXV2() 'reformat Medical History V2

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
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Filtered Medical History"
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
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
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

Sub RECISTV2_COHORT6() 'reformat recist_v2-cohort 6 into 3 tabs:taget lesions, non-target lesions, new lesions

Dim WS1 As Worksheet
Dim WS2 As Worksheet
Dim WS3 As Worksheet
Dim WS4 As Worksheet
Dim WB1 As Workbook
Dim RowNum As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer
Dim i2 As Integer
Dim j2 As Integer
Dim k2 As Integer
Dim l2 As Integer
Dim m2 As Integer
Dim i3 As Integer
Dim j3 As Integer
Dim k3 As Integer
Dim l3 As Integer
Dim m3 As Integer

Set WB1 = ActiveWorkbook
Set WS1 = ActiveSheet
'Add WS2 as output for taget lesion
With WB1
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Target Lesions"
    Set WS2 = Sheets("Target Lesions")
End With
'Set up header for WS2
WS1.Range("C8").Copy WS2.Range("A1")
For i = 27 To 30
    WS1.Cells(8, i).Copy WS2.Cells(1, i - 14) 'ws2 start from column 13 for repeat headers
Next i
'header for other common fields before the repeat items
WS2.Range("B1").Value = "Visit ID"
WS2.Range("C1").Value = "For Unscheduled Visits, Specify Day #"
WS2.Range("D1").Value = "Additional Infusion Number"
WS2.Range("E1").Value = "Days Post Infusion"
WS2.Range("F1").Value = "Timepoint"
WS2.Range("G1").Value = "Was RECIST Assessment Completed for this Timepoint?"
WS2.Range("H1").Value = "If no, specify reason"
WS2.Range("I1").Value = "Date of test"
WS2.Range("J1").Value = "Imaging Modality"
WS2.Range("K1").Value = "Other (Specify)"
WS2.Range("L1").Value = "Target Lesion Identified?"

'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
'Loop rows
l = 1
For i = 9 To RowNum
  If WS1.Cells(i, 21).Value = "Yes" Then 'RECIST done
  
    For j = 27 To 66 Step 4
        If WS1.Cells(i, j).Value <> "" Then
            l = l + 1
          'copy subject ID from worksheet 1 (9, 3) to worksheet 2 (2,1) all the way to worksheet 1(last row,3) to worksheet2(last row, 1)
                WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
                WS1.Cells(i, 16).Copy WS2.Cells(l, 2)
                WS1.Cells(i, 17).Copy WS2.Cells(l, 3)
                WS1.Cells(i, 18).Copy WS2.Cells(l, 4)
                WS1.Cells(i, 19).Copy WS2.Cells(l, 5)
                WS1.Cells(i, 20).Copy WS2.Cells(l, 6)
                WS1.Cells(i, 21).Copy WS2.Cells(l, 7)
                WS1.Cells(i, 22).Copy WS2.Cells(l, 8)
                WS1.Cells(i, 23).Copy WS2.Cells(l, 9)
                WS1.Cells(i, 24).Copy WS2.Cells(l, 10)
                WS1.Cells(i, 25).Copy WS2.Cells(l, 11)
                WS1.Cells(i, 26).Copy WS2.Cells(l, 12)
            For k = 0 To 3
                WS1.Cells(i, k + j).Copy WS2.Cells(l, k + 13)
            Next k
        End If
    Next j
Else 'RECIST not done
    l = l + 1
    WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
    WS1.Cells(i, 16).Copy WS2.Cells(l, 2)
    WS1.Cells(i, 17).Copy WS2.Cells(l, 3)
    WS1.Cells(i, 18).Copy WS2.Cells(l, 4)
    WS1.Cells(i, 19).Copy WS2.Cells(l, 5)
    WS1.Cells(i, 20).Copy WS2.Cells(l, 6)
    WS1.Cells(i, 21).Copy WS2.Cells(l, 7)
    WS1.Cells(i, 22).Copy WS2.Cells(l, 8)
    WS1.Cells(i, 23).Copy WS2.Cells(l, 9)
    WS1.Cells(i, 24).Copy WS2.Cells(l, 10)
    WS1.Cells(i, 25).Copy WS2.Cells(l, 11)
    WS1.Cells(i, 26).Copy WS2.Cells(l, 12)
  End If
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat

'Add WS3 as output for non-target lesion
With WB1
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Non-Target Lesions"
    Set WS3 = Sheets("Non-Target Lesions")
End With
'Set up header for WS3
WS1.Range("C8").Copy WS3.Range("A1")
For i2 = 68 To 72
    WS1.Cells(8, i2).Copy WS3.Cells(1, i2 - 55) 'ws3 start from column 13 for repeat headers
Next i2
'header for other common fields before the repeat items
WS3.Range("B1").Value = "Visit ID"
WS3.Range("C1").Value = "For Unscheduled Visits, Specify Day #"
WS3.Range("D1").Value = "Additional Infusion Number"
WS3.Range("E1").Value = "Days Post Infusion"
WS3.Range("F1").Value = "Timepoint"
WS3.Range("G1").Value = "Was RECIST Assessment Completed for this Timepoint?"
WS3.Range("H1").Value = "If no, specify reason"
WS3.Range("I1").Value = "Date of test"
WS3.Range("J1").Value = "Imaging Modality"
WS3.Range("K1").Value = "Other (Specify)"
WS3.Range("L1").Value = "Non-Target Lesion Identified?"

'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
'Loop rows
l2 = 1
For i2 = 9 To RowNum
  If WS1.Cells(i2, 21).Value = "Yes" Then 'RECIST done
  
    For j2 = 68 To 117 Step 5
        If WS1.Cells(i2, j2).Value <> "" Then
            l2 = l2 + 1
          'copy subject ID from worksheet 1 (9, 3) to worksheet 2 (2,1) all the way to worksheet 1(last row,3) to worksheet2(last row, 1)
                WS1.Cells(i2, 3).Copy WS3.Cells(l2, 1)
                WS1.Cells(i2, 16).Copy WS3.Cells(l2, 2)
                WS1.Cells(i2, 17).Copy WS3.Cells(l2, 3)
                WS1.Cells(i2, 18).Copy WS3.Cells(l2, 4)
                WS1.Cells(i2, 19).Copy WS3.Cells(l2, 5)
                WS1.Cells(i2, 20).Copy WS3.Cells(l2, 6)
                WS1.Cells(i2, 21).Copy WS3.Cells(l2, 7)
                WS1.Cells(i2, 22).Copy WS3.Cells(l2, 8)
                WS1.Cells(i2, 23).Copy WS3.Cells(l2, 9)
                WS1.Cells(i2, 24).Copy WS3.Cells(l2, 10)
                WS1.Cells(i2, 25).Copy WS3.Cells(l2, 11)
                WS1.Cells(i2, 67).Copy WS3.Cells(l2, 12)
            For k2 = 0 To 4
                WS1.Cells(i2, k2 + j2).Copy WS3.Cells(l2, k2 + 13)
            Next k2
        End If
    Next j2
Else 'RECIST not done
    l2 = l2 + 1
    WS1.Cells(i2, 3).Copy WS3.Cells(l2, 1)
    WS1.Cells(i2, 16).Copy WS3.Cells(l2, 2)
    WS1.Cells(i2, 17).Copy WS3.Cells(l2, 3)
    WS1.Cells(i2, 18).Copy WS3.Cells(l2, 4)
    WS1.Cells(i2, 19).Copy WS3.Cells(l2, 5)
    WS1.Cells(i2, 20).Copy WS3.Cells(l2, 6)
    WS1.Cells(i2, 21).Copy WS3.Cells(l2, 7)
    WS1.Cells(i2, 22).Copy WS3.Cells(l2, 8)
    WS1.Cells(i2, 23).Copy WS3.Cells(l2, 9)
    WS1.Cells(i2, 24).Copy WS3.Cells(l2, 10)
    WS1.Cells(i, 25).Copy WS3.Cells(l2, 11)
    WS1.Cells(i2, 67).Copy WS3.Cells(l2, 12)
  End If
Next i2

'Format the completed sheet
WS3.Activate
Call Main.OutFormat


With WB1
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "New Lesions"
    Set WS4 = Sheets("New Lesions")
End With
'Set up header for WS4
WS1.Range("C8").Copy WS4.Range("A1")
For i3 = 119 To 122
    WS1.Cells(8, i3).Copy WS4.Cells(1, i3 - 104) 'ws4 start from column 15 for repeat headers
Next i3
'header for other common fields before the repeat items
WS4.Range("B1").Value = "Visit ID"
WS4.Range("C1").Value = "For Unscheduled Visits, Specify Day #"
WS4.Range("D1").Value = "Additional Infusion Number"
WS4.Range("E1").Value = "Days Post Infusion"
WS4.Range("F1").Value = "Timepoint"
WS4.Range("G1").Value = "Was RECIST Assessment Completed for this Timepoint?"
WS4.Range("H1").Value = "If no, specify reason"
WS4.Range("I1").Value = "Date of test"
WS4.Range("J1").Value = "Imaging Modality"
WS4.Range("K1").Value = "Other (Specify)"
WS4.Range("L1").Value = "Total Target Tumor Burden (Sum of Lesion Diameters)"
WS4.Range("M1").Value = "Total Target Tumor Burden Percent Change"
WS4.Range("N1").Value = "New Lesion Identified?"

'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
'Loop rows
l3 = 1
For i3 = 9 To RowNum
  If WS1.Cells(i3, 21).Value = "Yes" Then 'RECIST done
    If WS1.Cells(i3, 119).Value <> "Yes" Then 'No new lession
    l3 = l3 + 1
                WS1.Cells(i3, 3).Copy WS4.Cells(l3, 1)
                WS1.Cells(i3, 16).Copy WS4.Cells(l3, 2)
                WS1.Cells(i3, 17).Copy WS4.Cells(l3, 3)
                WS1.Cells(i3, 18).Copy WS4.Cells(l3, 4)
                WS1.Cells(i3, 19).Copy WS4.Cells(l3, 5)
                WS1.Cells(i3, 20).Copy WS4.Cells(l3, 6)
                WS1.Cells(i3, 21).Copy WS4.Cells(l3, 7)
                WS1.Cells(i3, 22).Copy WS4.Cells(l3, 8)
                WS1.Cells(i3, 23).Copy WS4.Cells(l3, 9)
                WS1.Cells(i3, 24).Copy WS4.Cells(l3, 10)
                WS1.Cells(i3, 25).Copy WS4.Cells(l3, 11)
                WS1.Cells(i3, 118).Copy WS4.Cells(l3, 14)
    Else 'there is new lesion
      For j3 = 119 To 158 Step 4
        If WS1.Cells(i3, j3).Value <> "" Then
            l3 = l3 + 1
          'copy subject ID from worksheet 1 (9, 3) to worksheet 2 (2,1) all the way to worksheet 1(last row,3) to worksheet2(last row, 1)
                WS1.Cells(i3, 3).Copy WS4.Cells(l3, 1)
                WS1.Cells(i3, 16).Copy WS4.Cells(l3, 2)
                WS1.Cells(i3, 17).Copy WS4.Cells(l3, 3)
                WS1.Cells(i3, 18).Copy WS4.Cells(l3, 4)
                WS1.Cells(i3, 19).Copy WS4.Cells(l3, 5)
                WS1.Cells(i3, 20).Copy WS4.Cells(l3, 6)
                WS1.Cells(i3, 21).Copy WS4.Cells(l3, 7)
                WS1.Cells(i3, 22).Copy WS4.Cells(l3, 8)
                WS1.Cells(i3, 23).Copy WS4.Cells(l3, 9)
                WS1.Cells(i3, 24).Copy WS4.Cells(l3, 10)
                WS1.Cells(i3, 25).Copy WS4.Cells(l3, 11)
                WS1.Cells(i3, 118).Copy WS4.Cells(l3, 14)
            For k3 = 0 To 3
                WS1.Cells(i3, k3 + j3).Copy WS4.Cells(l3, k3 + 15)
            Next k3
        End If
    Next j3
    End If 'end if for new lesion
Else 'RECIST not done
    l3 = l3 + 1
    WS1.Cells(i3, 3).Copy WS4.Cells(l3, 1)
    WS1.Cells(i3, 16).Copy WS4.Cells(l3, 2)
    WS1.Cells(i3, 17).Copy WS4.Cells(l3, 3)
    WS1.Cells(i3, 18).Copy WS4.Cells(l3, 4)
    WS1.Cells(i3, 19).Copy WS4.Cells(l3, 5)
    WS1.Cells(i3, 20).Copy WS4.Cells(l3, 6)
    WS1.Cells(i3, 21).Copy WS4.Cells(l3, 7)
    WS1.Cells(i3, 22).Copy WS4.Cells(l3, 8)
    WS1.Cells(i3, 23).Copy WS4.Cells(l3, 9)
    WS1.Cells(i3, 24).Copy WS4.Cells(l3, 10)
    WS1.Cells(i, 25).Copy WS4.Cells(l3, 11)
    WS1.Cells(i3, 118).Copy WS4.Cells(l3, 14)
  End If
        WS1.Cells(i3, 159).Copy WS4.Cells(l3, 12)
        WS1.Cells(i3, 160).Copy WS4.Cells(l3, 13)
Next i3

'Format the completed sheet
WS4.Activate
Call Main.OutFormat
End Sub

Sub QFSR(WS1, WS2, WS3, WS4, LastRow)
    'Copy cohort 5-7 data to another tab
     WS1.Range("A1").AutoFilter Field:=3, Criteria1:=Array( _
        "02916 Subject Study Calendar - Cohort 5", _
        "02916 Subject Study Calendar - Cohort 6", _
        "02916 Subject Study Calendar - Cohort 7"), Operator:=xlFilterValues
   
    WS1.Range("A1:J" & LastRow).Select
    Selection.Copy
    
    Set WS2 = Sheets.Add
    WS2.Name = "Cohorts 5-7 Form Status"
    WS2.Paste
    Call Main.OutFormat

    Sheets("All Cohorts Form Status Report").Move Before:=Sheets(1)
   
    'Copy cohort 1-3 data to another tab
    Sheets("All Cohorts Form Status Report").Select
    WS1.Range("A1").AutoFilter Field:=3, Criteria1:="02916-Subject Calendar"
    WS1.Range("A1:J" & LastRow).Select
    Selection.Copy
    
    Set WS3 = Sheets.Add
    WS3.Name = "Cohorts 1-3 Form Status"
    WS3.Paste
    Call Main.OutFormat
 
    Sheets("All Cohorts Form Status Report").Move Before:=Sheets(1)

    Set WS4 = Sheets.Add
    WS4.Name = "Form Status Overview"
    
    WS4.Range("A1").Value = "All Cohorts Form Status"
    Call FormStatusOverview(WS1, WS4, 1, LastRow)
    
    WS4.Range("A8").Value = "Cohorts 1-3 Form Status"
    Call FormStatusOverview(WS3, WS4, 8, LastRow)
    
    WS4.Range("A15").Value = "Cohorts 5-7 Form Status"
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
WS2.Name = "Cohorts 5-7 Query Report"

Set WS3 = Sheets.Add(After:=WS2)
WS3.Name = "Cohorts 1-3 Query Report"

Set WS4 = Sheets.Add(Before:=WS1)
WS4.Name = "Query Report Overview"

WS1.Activate
WS1.Range("A1").Select
LastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
'Copy cohort 5-7 data to another tab
WS1.Range("A1").AutoFilter Field:=5, Criteria1:=Array( _
    "02916 Subject Study Calendar - Cohort 5", _
    "02916 Subject Study Calendar - Cohort 6", _
    "02916 Subject Study Calendar - Cohort 7"), Operator:=xlFilterValues
   
 WS1.Range("A1:S" & LastRow).SpecialCells(xlCellTypeVisible).Copy
 WS2.Paste

 'Copy cohort 1-3 data to another tab
 WS1.Range("A1").AutoFilter Field:=5, Criteria1:="02916-Subject Calendar"
 WS1.Range("A1:S" & LastRow).SpecialCells(xlCellTypeVisible).Copy
 WS3.Paste
 
 WS4.Range("A1").Value = "All Cohorts Query Status"
 Call QueryReportOverview(WS1, WS4, 1)
 
 WS4.Range("A7").Value = "Cohorts 1-3 Query Status"
 Call QueryReportOverview(WS3, WS4, 7)
 
 WS4.Range("A13").Value = "Cohorts 5-7 Query Status"
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
