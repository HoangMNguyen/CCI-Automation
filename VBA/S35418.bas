Attribute VB_Name = "S35418"
Sub BRIDGING_THERAPY()

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
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Filtered Bridging Therapy"
    Set WS2 = Sheets("Filtered Bridging Therapy")
End With
'Set up header for WS2
WS1.Range("C8").Copy WS2.Range("A1")
For i = 16 To 35
    WS1.Cells(8, i).Copy WS2.Cells(1, i - 14)
Next i


'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
'Loop rows
l = 1
For i = 9 To RowNum
    For j = 16 To 215 Step 20
    If WS1.Cells(i, j).Value <> "No" Then
        l = l + 1
        WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
        For k = 0 To 19
                WS1.Cells(i, k + j).Copy WS2.Cells(l, k + 2)
        Next k
    End If
    Next j
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat
End Sub
Sub LTFU_ANP()

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
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Filtered LTFU ANP"
    Set WS2 = Sheets("Filtered LTFU ANP")
End With
'Set up header for WS2
WS1.Range("C8").Copy WS2.Range("A1")
For i = 16 To 35
    WS1.Cells(8, i).Copy WS2.Cells(1, i - 14)
Next i



'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
'Loop rows
l = 1
For i = 9 To RowNum
    For j = 16 To 215 Step 20
    If WS1.Cells(i, j).Value <> "No" Then
        l = l + 1
        WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
        For k = 0 To 19
                WS1.Cells(i, k + j).Copy WS2.Cells(l, k + 2)
        Next k
    End If
    Next j
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat
End Sub
Sub PRIOR_ANP()

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
For i = 16 To 32
    WS1.Cells(8, i).Copy WS2.Cells(1, i - 14)
Next i



'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
'Loop rows
l = 1
For i = 9 To RowNum
    For j = 16 To 185 Step 17
    If WS1.Cells(i, j).Value <> "No" Then
        l = l + 1
        WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
        For k = 0 To 16
                WS1.Cells(i, k + j).Copy WS2.Cells(l, k + 2)
        Next k
    End If
    Next j
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat
End Sub


Sub MEDHX()

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

Sub PRIMARY_ANP()

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
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Filtered Primary ANP"
    Set WS2 = Sheets("Filtered Primary ANP")
End With
'Set up header for WS2
WS1.Range("C8").Copy WS2.Range("A1")
For i = 16 To 35 'medication/therapy to ongoing
    WS1.Cells(8, i).Copy WS2.Cells(1, i - 14)
Next i

'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row 'search backward from last row to the row has content
'Loop rows
l = 1
For i = 9 To RowNum
    'from column 16  (medication/therapy) to last column 215, step 20 (medication/therapy to ongoing has 20 columns)
    For j = 16 To 215 Step 20
      If WS1.Cells(i, j).Value <> "No" Then 'only copy the data where medication/therapy<>No
        l = l + 1
        'copy subject ID from worksheet 1 (9, 3) to worksheet 2 (2,1) all the way to worksheet 1(last row,3) to worksheet2(last row, 1)
        WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
        For k = 0 To 19
                WS1.Cells(i, k + j).Copy WS2.Cells(l, k + 2) 'copy data from worksheet 1 (9,16)-(9,35) to worksheet 2 (2, 2)-(2,21)
        Next k
     End If
    Next j
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat

End Sub

Sub PRIOR_TRANSPLANT_V2()

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
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Filtered Prior Transplant"
    Set WS2 = Sheets("Filtered Prior Transplant")
End With
'Set up header for WS2
WS1.Range("C8").Copy WS2.Range("A1") 'subject ID
For i = 20 To 28 'medication# to end date
    WS1.Cells(8, i).Copy WS2.Cells(1, i - 14) 'ws2 start to copy repeat headers from column F(6)
Next i
'header for other common fields before the repeat items
WS2.Range("B1").Value = "Did the subject receive a prior alloHCT?"
WS2.Range("C1").Value = "Date of transplant"
WS2.Range("D1").Value = "Donor Type"
WS2.Range("E1").Value = "Conditioning Type"

'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row 'search backward from last row to the row has content
'Loop rows
l = 1
For i = 9 To RowNum
If WS1.Cells(i, 16).Value <> "No" Then  'Did the subject receive a prior alloHCT<>No
    'from column 16  (medication#) to last column 64, step 9 (medication# to End Date has 9 columns)
    For j = 20 To 64 Step 9 'medication#1 to last column
        If WS1.Cells(i, j).Value <> "Not Applicable" Then 'only copy the data where medication<>Not applicable
            l = l + 1
            'copy subject ID from worksheet 1 (9, 3) to worksheet 2 (2,1) all the way to worksheet 1(last row,3) to worksheet2(last row, 1)
            WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
            WS1.Cells(i, 16).Copy WS2.Cells(l, 2)
            WS1.Cells(i, 17).Copy WS2.Cells(l, 3)
            WS1.Cells(i, 18).Copy WS2.Cells(l, 4)
            WS1.Cells(i, 19).Copy WS2.Cells(l, 5)
            For k = 0 To 8
                    WS1.Cells(i, k + j).Copy WS2.Cells(l, k + 6) 'copy data from worksheet 1 (9,20)-(9,28)(med# to end date) to worksheet 2 (2, 6)-(2,14), then next row all the way to the last row with data
            Next k
        End If
    Next j
Else 'Subject did not receive transplant, only copy subject ID and No
    l = l + 1
    WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
    WS1.Cells(i, 16).Copy WS2.Cells(l, 2)
End If
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat

End Sub


Sub TRANSPLANT_V2()

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
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Filtered Transplant"
    Set WS2 = Sheets("Filtered Transplant")
End With
'Set up header for WS2
WS1.Range("C8").Copy WS2.Range("A1") 'header for subject ID
For i = 21 To 29 'header for medication# to end date
   WS1.Cells(8, i).Copy WS2.Cells(1, i - 14)
Next i
'header for other common fields before the repeat items
WS2.Range("B1").Value = "Did the subject receive an alloHCT on this trial?"
WS2.Range("C1").Value = "Date of transplant"
WS2.Range("D1").Value = "Study Day"
WS2.Range("E1").Value = "Donor Type"
WS2.Range("F1").Value = "Conditioning Type"
WS2.Range("G1").Value = "Medication"

'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row 'search backward from last row to the row has content
'Loop rows
l = 1
For i = 9 To RowNum
If WS1.Cells(i, 16).Value <> "No" Then
    'from column 16  (medication#) to last column 70, step 9 (medication# to End Date has 9 columns)
  For j = 21 To 65 Step 9 'column 21 is medication #1
    If WS1.Cells(i, j).Value <> "Not Applicable" Then 'only copy the data where medication<>Not applicable
        l = l + 1
        'copy subject ID from worksheet 1 (9, 3) to worksheet 2 (2,1) all the way to worksheet 1(last row,3) to worksheet2(last row, 1)
        WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
        WS1.Cells(i, 16).Copy WS2.Cells(l, 2)
        WS1.Cells(i, 17).Copy WS2.Cells(l, 3)
        WS1.Cells(i, 18).Copy WS2.Cells(l, 4)
        WS1.Cells(i, 19).Copy WS2.Cells(l, 5)
        WS1.Cells(i, 20).Copy WS2.Cells(l, 6)
        WS1.Cells(i, 21).Copy WS2.Cells(l, 7)

        'For k = 0 To 9 'change to 8 for v2
        For k = 0 To 8
            WS1.Cells(i, k + j).Copy WS2.Cells(l, k + 7) 'copy data from worksheet 1 (9,21)to worksheet 2 (2, 7)...
          Next k
        End If
    Next j
Else 'Subject did not receive transplant, only copy subject ID and No
    l = l + 1
    WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
    WS1.Cells(i, 16).Copy WS2.Cells(l, 2)
End If
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat
End Sub


Sub TRANSFUSION()

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
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Filtered Transfusion"
    Set WS2 = Sheets("Filtered Transfusion")
End With
'Set up header for WS2
WS1.Range("C8").Copy WS2.Range("A1")
For i = 16 To 19 'Transfusion Data to Report to Total Number of Units Received
    WS1.Cells(8, i).Copy WS2.Cells(1, i - 14)
Next i


'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row 'search backward from last row to the row has content
'Loop rows
l = 1
For i = 9 To RowNum
    'from column 16  (transfustion data to report) to last column 35, step 4 (transfustion data to report to Total Number of Units Received has 4 columns)
    For j = 16 To 35 Step 4
    If WS1.Cells(i, j).Value <> "No" Then 'only copy the data where transfustion data to report<>No
        l = l + 1
        'copy subject ID from worksheet 1 (9, 3) to worksheet 2 (2,1) all the way to worksheet 1(last row,3) to worksheet2(last row, 1)
        WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
        For k = 0 To 3
            WS1.Cells(i, k + j).Copy WS2.Cells(l, k + 2) 'copy WS1(9,16) to WS2(2,2)..
          Next k
        End If
    Next j
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat
End Sub

Sub INFUSIONVITALS()

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
WS2.Range("C1").Value = "Unscheduled Day#"

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
