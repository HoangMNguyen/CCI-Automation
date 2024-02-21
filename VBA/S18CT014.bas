Attribute VB_Name = "S18CT014"
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

Application.ScreenUpdating = False


Set WB1 = ActiveWorkbook
Set WS1 = ActiveSheet
'Add WS2 as output
With WB1
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Filtered Prior ANP"
    Set WS2 = Sheets("Filtered Prior ANP")
End With
'Set up header for WS2
WS1.Range("C8").Copy WS2.Range("A1")
For i = 16 To 24
    WS1.Cells(8, i).Copy WS2.Cells(1, i - 14)
Next i



'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
'Loop rows
l = 1
For i = 9 To RowNum
    For j = 16 To 195 Step 9
    If WS1.Cells(i, j).Value <> "No" Then
        l = l + 1
        WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
        For k = 0 To 8
                WS1.Cells(i, k + j).Copy WS2.Cells(l, k + 2)
        Next k
    End If
    Next j
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat

Application.ScreenUpdating = True


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


Application.ScreenUpdating = False

Set WB1 = ActiveWorkbook
Set WS1 = ActiveSheet
'Add WS2 as output
With WB1
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Filtered Primary ANP"
    Set WS2 = Sheets("Filtered Primary ANP")
End With
'Set up header for WS2
WS1.Range("C8").Copy WS2.Range("A1")
For i = 16 To 26
    WS1.Cells(8, i).Copy WS2.Cells(1, i - 14)
Next i



'Count number of rows w/ header
WS1.Activate
WS1.Range("A1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
'Loop rows
l = 1
For i = 9 To RowNum
    m = 1
    For j = 16 To 125 Step 11
    If (WS1.Cells(i, j).Value <> "No") Or (m = 10) Then
        l = l + 1
        WS1.Cells(i, 3).Copy WS2.Cells(l, 1)
        For k = 0 To 10
                WS1.Cells(i, k + j).Copy WS2.Cells(l, k + 2)
        Next k
    Else
        m = m + 1
    End If
    Next j
Next i

'Format the completed sheet
WS2.Activate
Call Main.OutFormat

Application.ScreenUpdating = True


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

Application.ScreenUpdating = False

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

Application.ScreenUpdating = True

End Sub

