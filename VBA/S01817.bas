Attribute VB_Name = "S01817"
Sub ConMedLog()

Dim WS1 As Worksheet
Dim WS2 As Worksheet
Dim WB1 As Workbook
Dim RowNum As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer
Dim SubjectID As String

Set WB1 = ActiveWorkbook
Set WS1 = ActiveSheet
ActiveSheet.Cells.UnMerge
'Add WS2 as output
With WB1
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Filtered ConMed"
    Set WS2 = Sheets("Filtered ConMed")
End With
'Set up header for WS2
WS1.Range("C12").Copy WS2.Range("A1") 'header is in row 11
For i = 17 To 33
    WS1.Cells(12, i).Copy WS2.Cells(1, i - 14)
Next i
'header for subject #ID
WS2.Range("A1").Value = "Subject ID"
WS2.Range("B1").Value = "Medication#"


'Count number of rows w/ header
WS1.Activate
WS1.Range("P1").Select
RowNum = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row 'search backward from last row to the row has content
'Loop rows
'Assign first locol variable
l = 0
SubjectID = WS1.Cells(12, 3).Value
For i = 12 To RowNum
    If WS1.Range("P" & i).Value = "Medication" Then
        SubjectID = WS1.Cells(i, 3).Value
        l = l + 1
    ElseIf WS1.Range("P" & i).Value <> "Medication" And WS1.Range("P" & i).Value <> "-" Then
        WS2.Cells(i - 10 - l, 1).Value = SubjectID
        WS1.Range("P" & i & ":" & "AG" & i).Copy WS2.Range("B" & (i - 10 - l) & ":" & "S" & (i - 10 - l))
    ElseIf WS1.Range("P" & i).Value = "-" Then
        l = l + 1
    End If
Next i

'Format the completed sheet
WS2.Activate
ActiveSheet.Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Font.Bold = True
Range(Selection, Selection.End(xlDown)).Select
Selection.WrapText = False
Selection.Columns.AutoFit
Selection.VerticalAlignment = xlCenter
Selection.HorizontalAlignment = xlLeft
Selection.Borders.LineStyle = xlContinuous
Selection.Borders.Weight = xlThin
Selection.AutoFilter

End Sub



