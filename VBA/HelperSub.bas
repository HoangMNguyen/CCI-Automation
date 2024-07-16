Attribute VB_Name = "HelperSub"
Option Explicit

Public Sub RemoveFilter()

    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If

End Sub



Public Sub FormatTable() 'bold selected row, autofit columns, add border
    If ActiveSheet.Range(Selection.Address).Offset(0, 1) <> 0 And ActiveSheet.Range(Selection.Address).Offset(1, 0) <> 0 Then
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Font.Bold = True
        
        Range(Selection, Selection.End(xlDown)).Select
        Selection.WrapText = False
        Selection.Columns.AutoFit
        Selection.VerticalAlignment = xlCenter
        Selection.HorizontalAlignment = xlLeft
        Selection.Borders.LineStyle = xlContinuous
        Selection.Borders.Weight = xlThin
        Selection.Font.Size = 10
        Selection.Borders.LineStyle = xlContinuous
        Selection.Borders.Color = vbBlack
        Selection.Borders.Weight = xlThin
    ElseIf ActiveSheet.Range(Selection.Address).Offset(0, 1) = 0 And ActiveSheet.Range(Selection.Address).Offset(1, 0) <> 0 Then
        Selection.Font.Bold = True
        
        Range(Selection, Selection.End(xlDown)).Select
        Selection.WrapText = False
        Selection.Columns.AutoFit
        Selection.VerticalAlignment = xlCenter
        Selection.HorizontalAlignment = xlLeft
        Selection.Borders.LineStyle = xlContinuous
        Selection.Borders.Weight = xlThin
        Selection.Font.Size = 10
        Selection.Borders.LineStyle = xlContinuous
        Selection.Borders.Color = vbBlack
        Selection.Borders.Weight = xlThin
        
    ElseIf ActiveSheet.Range(Selection.Address).Offset(0, 1) <> 0 And ActiveSheet.Range(Selection.Address).Offset(1, 0) = 0 Then
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Font.Bold = True
        Selection.WrapText = False
        Selection.Columns.AutoFit
        Selection.VerticalAlignment = xlCenter
        Selection.HorizontalAlignment = xlLeft
        Selection.Borders.LineStyle = xlContinuous
        Selection.Borders.Weight = xlThin
        Selection.Font.Size = 10
        Selection.Borders.LineStyle = xlContinuous
        Selection.Borders.Color = vbBlack
        Selection.Borders.Weight = xlThin
    End If
    

End Sub

Public Sub FormatTable2() 'bold selected row, autofit columns, add border +  keeping wrap text
    If ActiveSheet.Range(Selection.Address).Offset(0, 1) <> 0 And ActiveSheet.Range(Selection.Address).Offset(1, 0) <> 0 Then
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Font.Bold = True
        
        Range(Selection, Selection.End(xlDown)).Select
        Selection.WrapText = True
        Selection.Columns.AutoFit
        Selection.VerticalAlignment = xlCenter
        Selection.HorizontalAlignment = xlCenter
        Selection.Borders.LineStyle = xlContinuous
        Selection.Borders.Weight = xlThin
        Selection.Font.Size = 10
        Selection.Borders.LineStyle = xlContinuous
        Selection.Borders.Color = vbBlack
        Selection.Borders.Weight = xlThin
    ElseIf ActiveSheet.Range(Selection.Address).Offset(0, 1) = 0 And ActiveSheet.Range(Selection.Address).Offset(1, 0) <> 0 Then
        Selection.Font.Bold = True
        
        Range(Selection, Selection.End(xlDown)).Select
        Selection.WrapText = True
        Selection.Columns.AutoFit
        Selection.VerticalAlignment = xlCenter
        Selection.HorizontalAlignment = xlCenter
        Selection.Borders.LineStyle = xlContinuous
        Selection.Borders.Weight = xlThin
        Selection.Font.Size = 10
        Selection.Borders.LineStyle = xlContinuous
        Selection.Borders.Color = vbBlack
        Selection.Borders.Weight = xlThin
        
    ElseIf ActiveSheet.Range(Selection.Address).Offset(0, 1) <> 0 And ActiveSheet.Range(Selection.Address).Offset(1, 0) = 0 Then
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Font.Bold = True
        Selection.WrapText = True
        Selection.Columns.AutoFit
        Selection.VerticalAlignment = xlCenter
        Selection.HorizontalAlignment = xlCenter
        Selection.Borders.LineStyle = xlContinuous
        Selection.Borders.Weight = xlThin
        Selection.Font.Size = 10
        Selection.Borders.LineStyle = xlContinuous
        Selection.Borders.Color = vbBlack
        Selection.Borders.Weight = xlThin
    End If
    

End Sub

Public Function duplicateWorkbook(wk As Workbook) As Workbook
    Dim path As String
    path = Environ("temp") & "\" & wk.Name & "." & _
        Right(wk.FullName, Len(wk.FullName) - InStrRev(wk.FullName, "."))
    wk.SaveCopyAs path
    Set duplicateWorkbook = Workbooks.Add(path)
    Kill path
End Function

Public Function getLastModified() As Date
   getLastModified = ActiveWorkbook.BuiltinDocumentProperties("Last Save Time")
End Function

Public Function Now2Date(mDate As Date)
    Now2Date = Format(mDate, "YYMMDD")
End Function

Public Function Now2Time(mDate As Date)
    Now2Time = Format(mDate, "hmm")
End Function

Public Function showAll()

If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
  ActiveSheet.ShowAllData
End If

End Function

Public Function sortAsc(colHeaderAddress)
With ActiveSheet.Sort
    .SortFields.Add Key:=Range(colHeaderAddress), Order:=xlAscending
    .SetRange Selection
    .header = xlYes
    .Apply
End With
End Function

Public Function sortDes(colHeaderAddress)
With ActiveSheet.Sort
    .SortFields.Add Key:=Range(colHeaderAddress), Order:=xlDescending
    .SetRange Selection
    .header = xlYes
    .Apply
End With
End Function

Public Function GetColumnLetter(colNum As Long) As String
    Dim vArr
    vArr = Split(Cells(1, colNum).Address(True, False), "$")
    GetColumnLetter = vArr(0)
End Function

Public Function ImportSheet(SheetNum As Long)
Dim sImportFile As String, sFile As String
Dim sThisBk As Workbook
Dim vfilename As Variant
Dim wbBk As Workbook
Dim wsSht As Worksheet


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Set sThisBk = ActiveWorkbook
sImportFile = Application.GetOpenFilename( _
FileFilter:="CSV Files (.csv), *.csv", Title:="Open Workbook")
If sImportFile = "False" Then
MsgBox "No File Selected!"
Exit Function

Else
vfilename = Split(sImportFile, "\")
sFile = vfilename(UBound(vfilename))
Application.Workbooks.Open fileName:=sImportFile

Set wbBk = Workbooks(sFile)
With wbBk
Set wsSht = .Sheets(1)
wsSht.Copy Before:=sThisBk.Sheets(SheetNum)
wbBk.Close SaveChanges:=False
End With
End If
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Function

Public Function ImportSheetByName(FilePath As String, fileName As String)
Dim sImportFile As String, sFile As String
Dim sThisBk As Workbook
Dim vfilename As Variant
Dim wbBk As Workbook
Dim wsSht As Worksheet


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Set sThisBk = ActiveWorkbook

vfilename = Split(sImportFile, "\")

Application.Workbooks.Open fileName:=FilePath & fileName

Set wbBk = Workbooks(Workbooks.Count)
With wbBk
Set wsSht = .Sheets(1)
wsSht.Copy Before:=sThisBk.Sheets(Sheets.Count)
wbBk.Close SaveChanges:=False
End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Function

Public Sub CopyUnique(CopyWS, CopyR, PasteWS, PasteR)

With CopyWS
      .Range(CopyR).Copy
End With
With PasteWS
    .Range(PasteR).PasteSpecial xlPasteValues
    .Range(PasteR, .Range(PasteR).End(xlDown)).RemoveDuplicates 1, xlNo
    .Range(PasteR, .Range(PasteR).End(xlDown)).Sort Key1:=.Range(PasteR), Order1:=xlAscending, header:=xlNo
End With
End Sub
Public Sub CopyColumn(CopyWS, CopyR, PasteWS, PasteR)

With CopyWS
      .Range(CopyR).Copy
End With
With PasteWS
    .Range(PasteR).PasteSpecial xlPasteValues
    .Range(PasteR, .Range(PasteR).End(xlDown)).Sort Key1:=.Range(PasteR), Order1:=xlAscending, header:=xlNo
End With
End Sub

Public Function CountPerColumn(CountSheet, ColumnNum, Criterial, Optional ColumnNum2 As Long = 0, Optional Criterial2 As String = "none")
CountSheet.Activate
RemoveFilter

If ColumnNum2 = 0 Then
CountSheet.Range("A1").AutoFilter Field:=ColumnNum, Criteria1:=Criterial, Operator:=xlFilterValues
CountPerColumn = CountSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
Else
With CountSheet.Range("A1")
    .AutoFilter Field:=ColumnNum, Criteria1:=Criterial
    .AutoFilter Field:=ColumnNum2, Criteria1:=Criterial2
CountPerColumn = CountSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
End With
End If

End Function

Public Function CountToPercent(CountNum, Total)

CountToPercent = "(" & Trim(Str(Round(CountNum / Total * 100, 2))) & "%)"

End Function


Public Function FindLastRowA(Sheet As Worksheet)

'only find the last row of the sheet assuming the sheet has column A data

Sheet.Activate
Sheet.Range("A1").Select
FindLastRowA = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row

End Function

Public Function SelectFolder()
'PURPOSE: Have User Select a Folder Path and Store it to a variable
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim FldrPicker As FileDialog
Dim myFolder As String

'Have User Select Folder to Save to with Dialog Box
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

  With FldrPicker
    .Title = "Select A Target Folder"
    .AllowMultiSelect = False
    If .Show <> -1 Then Exit Function 'Check if user clicked cancel button
    myFolder = .SelectedItems(1) & "\"
  End With
  
'Carry out rest of your code here....
SelectFolder = myFolder

End Function
Public Function L2N(ColumnLetter As String)

Dim ColumnNumber As Long

'Convert To Column Number
ColumnNumber = Range(ColumnLetter & 1).Column
  
L2N = ColumnNumber
    
End Function
Public Function N2L(ColumnNumber As Long)

Dim ColumnLetter As String

'Convert To Column Letter
  ColumnLetter = Split(Cells(1, ColumnNumber).Address, "$")(1)
  
N2L = ColumnLetter
  
End Function

Public Function RemoveColumn(Worksheet As Worksheet, ColumnName As String)

    ' Activate the intended worksheet
    Worksheet.Activate

    ' Declare a variable to represent the column range
    Dim TargetColumn As Range

    ' Search the active worksheet for the column name
    ' The search is performed in the first row (header row) of the worksheet
    Set TargetColumn = Worksheet.Rows(1).Find(What:=ColumnName, LookIn:=xlValues, LookAt:=xlWhole)

    ' Check if the column was found
    If Not TargetColumn Is Nothing Then
        ' Find the index of the column
        Dim columnIndex As Integer
        columnIndex = TargetColumn.Column

        ' Remove the column with that index
        Worksheet.Columns(columnIndex).Delete
    End If

End Function

Public Function SetWidthWrapColumn(Worksheet As Worksheet, ColumnName As String, Width As Integer)

    ' Activate the intended worksheet
    Worksheet.Activate

    ' Declare a variable to represent the column range
    Dim TargetColumn As Range

    ' Search the active worksheet for the column name
    ' The search is performed in the first row (header row) of the worksheet
    Set TargetColumn = Worksheet.Rows(1).Find(What:=ColumnName, LookIn:=xlValues, LookAt:=xlWhole)

    ' Check if the column was found
    If Not TargetColumn Is Nothing Then
        ' Find the index of the column
        Dim columnIndex As Integer
        columnIndex = TargetColumn.Column

        ' Set the width of the column and wrap
        Worksheet.Columns(columnIndex).ColumnWidth = Width
        Worksheet.Columns(columnIndex).WrapText = True
    End If

End Function

Public Function FilterColumn(Worksheet As Worksheet, ColumnName As String, Optional FilterValues As Variant)

    ' Activate the intended worksheet
    Worksheet.Activate

    ' Declare a variable to represent the column range
    Dim TargetColumn As Range

    ' Search the active worksheet for the column name
    Set TargetColumn = Worksheet.Rows(1).Find(What:=ColumnName, LookIn:=xlValues, LookAt:=xlWhole)

    ' Check if the column was found
    If Not TargetColumn Is Nothing Then
        Dim columnIndex As Integer
        columnIndex = TargetColumn.Column

        ' Clear existing filters
        'If Worksheet.AutoFilterMode Then Worksheet.AutoFilterMode = False

        ' Apply the filter based on provided values
        If IsEmpty(FilterValues) Then
            ' Show all data if no filter is provided
            Worksheet.Range("A1").AutoFilter
        Else
            Worksheet.Range("A1").AutoFilter Field:=columnIndex, Criteria1:=FilterValues, Operator:=xlFilterValues
        End If
    End If

End Function

Public Function CountPerColumnName(CountSheet As Worksheet, ColumnName1 As String, Criterial As Variant, _
                                   Optional ColumnName2 As String = "", Optional Criterial2 As Variant = "none") As Long

    Dim ColumnNum As Integer
    Dim ColumnNum2 As Integer
    Dim TargetColumn As Range

    ' Activate the worksheet
    CountSheet.Activate

    ' Remove any existing filters
    If CountSheet.AutoFilterMode Then CountSheet.AutoFilterMode = False

    ' Find the first column index
    Set TargetColumn = CountSheet.Rows(1).Find(What:=ColumnName1, LookIn:=xlValues, LookAt:=xlWhole)
    If Not TargetColumn Is Nothing Then
        ColumnNum = TargetColumn.Column
    Else
        Exit Function
    End If

    ' Check if a second column name is provided and find its index
    If Len(ColumnName2) > 0 Then
        Set TargetColumn = CountSheet.Rows(1).Find(What:=ColumnName2, LookIn:=xlValues, LookAt:=xlWhole)
        If Not TargetColumn Is Nothing Then
            ColumnNum2 = TargetColumn.Column
        Else
            ' If ColumnName2 is provided but not found, exit the function
            Exit Function
        End If
    End If

    ' Apply filters and count
    If Len(ColumnName2) = 0 Then
        CountSheet.Range("A1").AutoFilter Field:=ColumnNum, Criteria1:=Criterial
        CountPerColumnName = CountSheet.AutoFilter.Range.Columns(ColumnNum).SpecialCells(xlCellTypeVisible).Cells.Count - 1
    Else
        With CountSheet.Range("A1")
            .AutoFilter Field:=ColumnNum, Criteria1:=Criterial
            .AutoFilter Field:=ColumnNum2, Criteria1:=Criterial2
            CountPerColumnName = CountSheet.AutoFilter.Range.Columns(ColumnNum).SpecialCells(xlCellTypeVisible).Cells.Count - 1
        End With
    End If

    ' Remove filters after counting
    If CountSheet.AutoFilterMode Then CountSheet.AutoFilterMode = False

End Function


' Helper function to find column number
Private Function GetColumnNumber(ws As Worksheet, colName As String) As Integer
    Dim col As Range
    Set col = ws.Rows(1).Find(What:=colName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not col Is Nothing Then
        GetColumnNumber = col.Column
    Else
        GetColumnNumber = 0
    End If
End Function

Public Function FilterExcludeValue(Worksheet As Worksheet, ColumnName As String, Optional FilterValues As Variant)

    ' Activate the intended worksheet
    Worksheet.Activate

    ' Declare a variable to represent the column range
    Dim TargetColumn As Range

    ' Search the active worksheet for the column name
    Set TargetColumn = Worksheet.Rows(1).Find(What:=ColumnName, LookIn:=xlValues, LookAt:=xlWhole)

    ' Check if the column was found
    If Not TargetColumn Is Nothing Then
        Dim columnIndex As Integer
        columnIndex = TargetColumn.Column

        ' Clear existing filters
        'If Worksheet.AutoFilterMode Then Worksheet.AutoFilterMode = False

        ' Apply the filter based on provided values
        If IsEmpty(FilterValues) Then
            ' Show all data if no filter is provided
            Worksheet.Range("A1").AutoFilter
        Else
            Worksheet.Range("A1").AutoFilter Field:=columnIndex, Criteria1:="<>" & FilterValues
        End If
    End If
    
End Function


Public Function CountFilteredRows(Worksheet As Worksheet, ColumnName As String) As Long
    Dim TargetColumn As Range
    Dim lastRow As Long
    Dim VisibleCells As Range

    ' Check if there's an active AutoFilter on the worksheet
    If Not Worksheet.AutoFilterMode Then
        MsgBox "No filter applied on the sheet."
        CountFilteredRows = 0
        Exit Function
    End If

    ' Find the target column
    Set TargetColumn = Worksheet.Rows(1).Find(What:=ColumnName, LookIn:=xlValues, LookAt:=xlWhole)
    If TargetColumn Is Nothing Then
        MsgBox "Column '" & ColumnName & "' not found."
        CountFilteredRows = 0
        Exit Function
    End If

    ' Find the last row with data in the target column
    lastRow = Worksheet.Cells(Worksheet.Rows.Count, TargetColumn.Column).End(xlUp).Row
    
    ' Check if there are no data rows other than the header
    If lastRow = 1 Then
        CountFilteredRows = 0
        Exit Function
    End If

    ' Define the range to count (excluding header)
    Set VisibleCells = Worksheet.Range(TargetColumn.Offset(1, 0), Worksheet.Cells(lastRow, TargetColumn.Column)).SpecialCells(xlCellTypeVisible)

    ' Count the visible cells in the range
    CountFilteredRows = VisibleCells.Cells.Count
End Function

Public Function CopySelectedVisibleColumnsToLocation(SourceWorksheet As Worksheet, DestinationWorksheet As Worksheet, _
                                              HeaderNames As Variant, DestColumn As Integer, DestRow As Integer)

    Dim sourceRange As Range, ColRange As Range
    Dim lastRow As Long, i As Integer, ColIndex As Integer
    Dim DestinationCell As Range

    ' Check if there's an active AutoFilter on the source worksheet
    If Not SourceWorksheet.AutoFilterMode Then
        MsgBox "No filter applied on the source sheet."
        Exit Function
    End If

    ' Find the last row with data in the source worksheet
    lastRow = SourceWorksheet.Cells(SourceWorksheet.Rows.Count, 1).End(xlUp).Row
    
    'exit if lastrow is 1
    If lastRow = 1 Then
        Exit Function
    End If

    ' Loop through the array of header names
    For i = LBound(HeaderNames) To UBound(HeaderNames)
        ' Find the column for each header name
        Set ColRange = SourceWorksheet.Rows(1).Find(What:=HeaderNames(i), LookIn:=xlValues, LookAt:=xlWhole)
        If Not ColRange Is Nothing Then
            ColIndex = ColRange.Column

            ' Define the range to copy for this column, excluding headers
            Set sourceRange = SourceWorksheet.Range(SourceWorksheet.Cells(2, ColIndex), SourceWorksheet.Cells(lastRow, ColIndex)).SpecialCells(xlCellTypeVisible)

            ' Set the destination range
            Set DestinationCell = DestinationWorksheet.Cells(DestRow, DestColumn + i - LBound(HeaderNames))

            ' Copy the visible cells of this column
            sourceRange.Copy DestinationCell

        Else
            MsgBox "Column with header '" & HeaderNames(i) & "' not found."
        End If
    Next i

    ' Optional: Clear clipboard to release memory
    Application.CutCopyMode = False
End Function

Public Function CopyColumnsWithHeaders(SourceWorksheet As Worksheet, DestinationWorksheet As Worksheet, _
                                              HeaderNames As Variant, DestColumn As Integer, DestRow As Integer)

    Dim sourceRange As Range, ColRange As Range
    Dim lastRow As Long, i As Integer, ColIndex As Integer
    Dim DestinationCell As Range

    ' Ensure HeaderNames is an array
    If Not IsArray(HeaderNames) Then
        MsgBox "HeaderNames must be an array."
        
        Exit Function
    End If

    ' Find the last row with data in the source worksheet
    lastRow = SourceWorksheet.Cells(SourceWorksheet.Rows.Count, 1).End(xlUp).Row
    
    'if lastrow is 1
    If lastRow = 1 Then
        ' Loop through the array of header names
        For i = LBound(HeaderNames) To UBound(HeaderNames)
            ' Find the column for each header name
            Set ColRange = SourceWorksheet.Rows(1).Find(What:=HeaderNames(i), LookIn:=xlValues, LookAt:=xlWhole)
            If Not ColRange Is Nothing Then
                ' Define the range to copy for this column, including headers
                Set sourceRange = ColRange
    
                ' Set the destination range
                Set DestinationCell = DestinationWorksheet.Cells(DestRow, DestColumn + i - LBound(HeaderNames))
    
                ' Copy the visible cells of this column
                sourceRange.Copy DestinationCell
            End If
        Next i
    Else
        ' Loop through the array of header names
        For i = LBound(HeaderNames) To UBound(HeaderNames)
            ' Find the column for each header name
            Set ColRange = SourceWorksheet.Rows(1).Find(What:=HeaderNames(i), LookIn:=xlValues, LookAt:=xlWhole)
            If Not ColRange Is Nothing Then
                ColIndex = ColRange.Column
    
                ' Define the range to copy for this column, including headers
                Set sourceRange = SourceWorksheet.Range(SourceWorksheet.Cells(1, ColIndex), SourceWorksheet.Cells(lastRow, ColIndex))
    
                ' Set the destination range
                Set DestinationCell = DestinationWorksheet.Cells(DestRow, DestColumn + i - LBound(HeaderNames))
    
                ' Copy the visible cells of this column
                sourceRange.Copy DestinationCell
    
            Else
                MsgBox "Column with header '" & HeaderNames(i) & "' not found."
                Debug.Print "Column with header '" & HeaderNames(i) & "' not found."
            End If
        Next i
    End If

    ' Optional: Clear clipboard to release memory
    Application.CutCopyMode = False
End Function

Public Function ConvertHeadersToIndex(SourceWorksheet As Worksheet, HeaderNames As Variant) As Variant

    Dim headerIndexes() As Integer
    Dim header As Variant
    Dim columnIndex As Integer
    Dim i As Integer

    ReDim headerIndexes(LBound(headers) To UBound(headers))

    i = LBound(headers)
    For Each header In headers
        columnIndex = 0
        On Error Resume Next ' In case the header is not found
        columnIndex = Application.Match(header, ws.Rows(1), 0)
        On Error GoTo 0 ' Resume normal error handling

        If columnIndex > 0 Then
            headerIndexes(i) = columnIndex
        Else
            headerIndexes(i) = -1 ' or some other value indicating not found
        End If

        i = i + 1
    Next header

    ConvertHeadersToIndex = headerIndexes
    ' Optional: Clear clipboard to release memory
    Application.CutCopyMode = False
End Function

Public Function FindLastColumn(ws As Worksheet) As Long
    FindLastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
End Function

Public Function IsNumberInArray(num As Long, arr As Variant) As Boolean
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    Dim element As Variant

    ' Check if the input is actually an array
    If Not IsArray(arr) Then
        IsNumberInArray = False
        Exit Function
    End If

    ' Add array elements to the dictionary as keys
    For Each element In arr
        If Not dict.Exists(element) Then
            dict.Add element, Nothing
        End If
    Next element

    ' Check if the number exists in the dictionary
    IsNumberInArray = dict.Exists(num)
End Function

Public Function CopyColumnsWithIndex(SourceWorksheet As Worksheet, DestinationWorksheet As Worksheet, HeaderIndex As Long, DestColumn As Integer)

    Dim sourceRange As Range, ColRange As Range
    Dim lastRow As Long, i As Integer, ColIndex As Integer
    Dim DestinationCell As Range

    ' Find the last row with data in the source worksheet
    lastRow = SourceWorksheet.Cells(SourceWorksheet.Rows.Count, 1).End(xlUp).Row

    ' Define the range to copy for this column, including headers
    Set sourceRange = SourceWorksheet.Range(SourceWorksheet.Cells(1, HeaderIndex), SourceWorksheet.Cells(lastRow, HeaderIndex))

    ' Set the destination range
    Set DestinationCell = DestinationWorksheet.Cells(1, DestColumn)

    ' Copy the visible cells of this column
    sourceRange.Copy DestinationCell

    ' Optional: Clear clipboard to release memory
    Application.CutCopyMode = False
End Function

Public Function FindHeaderIndexes(ws As Worksheet, headers As Variant) As Variant
    Dim headerIndexes() As Integer
    Dim header As Variant
    Dim columnIndex As Integer
    Dim i As Integer

    ReDim headerIndexes(LBound(headers) To UBound(headers))

    i = LBound(headers)
    For Each header In headers
        columnIndex = 0
        On Error Resume Next ' In case the header is not found
        columnIndex = Application.Match(header, ws.Rows(1), 0)
        On Error GoTo 0 ' Resume normal error handling

        If columnIndex > 0 Then
            headerIndexes(i) = columnIndex
        Else
            headerIndexes(i) = -1 ' or some other value indicating not found
        End If

        i = i + 1
    Next header

    FindHeaderIndexes = headerIndexes
End Function

Public Function WarningCSV()

    'Check if the file extension is .csv
    Dim fileName As String
    Dim fileExtension As String

    ' Get the name of the active workbook
    fileName = ActiveWorkbook.Name

    ' Extract the file extension
    fileExtension = Right(fileName, Len(fileName) - InStrRev(fileName, "."))

    ' Check if the file extension is csv
    If LCase(fileExtension) = "csv" Then
        MsgBox "The active workbook is a CSV file. Please make sure to save it in Excel file format :)"
    End If
    
End Function

Function RoundDownColumn(ws As Worksheet, columnHeader As String)
    Dim headerRow As Long
    Dim headerCol As Long
    Dim lastRow As Long
    Dim cell As Range
    Dim found As Range
    
    headerRow = 1 ' Assuming the headers are in the first row
    Set found = ws.Rows(headerRow).Find(columnHeader, LookIn:=xlValues, LookAt:=xlWhole)

    ' Check if the column header is found
    If Not found Is Nothing Then
        headerCol = found.Column
        lastRow = ws.Cells(ws.Rows.Count, headerCol).End(xlUp).Row

        ' Loop through each cell in the column and round down
        For Each cell In ws.Range(ws.Cells(headerRow + 1, headerCol), ws.Cells(lastRow, headerCol))
            If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
                cell.Value = Application.WorksheetFunction.Floor(cell.Value, 1)
            End If
        Next cell
    End If
End Function

Function FindColumn(ws As Worksheet, headerName As String) As String
'Function find column header and return letter
'Last updated: 12/19/2023 By Hoang Nguyen
    Dim rng As Range
    Dim ColumnLetter As String

    ' Set the range to search as the first row of the worksheet
    Set rng = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole)

    ' Check if anything was found
    If Not rng Is Nothing Then
        ColumnLetter = Split(rng.Address, "$")(1)  ' Extracts the column letter
        FindColumn = ColumnLetter
    Else
        FindColumn = "Not Found"
    End If
End Function
Function ExtractUniqueValues(WS1 As Worksheet, ColumnNum As Long)
    Dim ws As Worksheet
    Set ws = WS1

    Dim sourceRange As Range
    Dim tempRange As Range
    Dim lastRow As Long
    Dim i As Long
    Dim lastPosition As Integer
    Dim ColumnLetter As String
    ' Set the range from which to extract unique values
    lastRow = FindLastRowA(ws)
    ColumnLetter = GetColumnLetter(ColumnNum)
    Set sourceRange = ws.Range(ColumnLetter & "1:" & ColumnLetter & lastRow)

    ' Create a temporary column to store the study number values
    Set tempRange = ws.Range("Z1:Z" & lastRow) ' Using column Z as temporary, change if needed
    For i = 1 To lastRow
        lastPosition = InStrRev(sourceRange.Cells(i, 1).Value, "-")
        If lastPosition > 0 Then
            tempRange.Cells(i, 1).Value = Left(sourceRange.Cells(i, 1).Value, lastPosition - 1)
        Else
            tempRange.Cells(i, 1).Value = sourceRange.Cells(i, 1).Value ' If the character is not found, return the whole string
        End If
    Next i

    ' Get unique values from the temporary column
    Dim uniqueValues As Variant
    uniqueValues = Application.WorksheetFunction.Unique(ws.Range("Z2:Z" & lastRow))

    ' Clear the temporary column
    tempRange.ClearContents

    ExtractUniqueValues = uniqueValues
    ' Outputting to Immediate Window for demonstration
    'For i = LBound(uniqueValues, 1) To UBound(uniqueValues, 1)
        'Debug.Print uniqueValues(i, 1)
    'Next i
End Function

Function DeepCopyArray(sourceArray As Variant) As Variant
    Dim copiedArray As Variant
    Dim i As Long, j As Long, k As Long
    Dim numDims As Integer

    ' Check if the source is actually an array
    If Not IsArray(sourceArray) Then
        Err.Raise Number:=vbObjectError + 513, _
                  Description:="Source is not an array."
        Exit Function
    End If

    ' Check if the array is empty
    If IsEmpty(sourceArray) Then
        DeepCopyArray = Array() ' Return an empty array if source is empty
        Exit Function
    End If

    ' Determine the number of dimensions
    On Error Resume Next
    numDims = 0
    Do While Err.Number = 0
        numDims = numDims + 1
        Dim test As Variant
        test = LBound(sourceArray, numDims)
    Loop
    On Error GoTo 0

    ' Copy the array based on the number of dimensions
    Select Case numDims
        Case 1 ' One-dimensional array
            ReDim copiedArray(LBound(sourceArray) To UBound(sourceArray))
            For i = LBound(sourceArray) To UBound(sourceArray)
                copiedArray(i) = sourceArray(i)
            Next i

        Case 2 ' Two-dimensional array
            ReDim copiedArray(LBound(sourceArray, 1) To UBound(sourceArray, 1), _
                              LBound(sourceArray, 2) To UBound(sourceArray, 2))
            For i = LBound(sourceArray, 1) To UBound(sourceArray, 1)
                For j = LBound(sourceArray, 2) To UBound(sourceArray, 2)
                    copiedArray(i, j) = sourceArray(i, j)
                Next j
            Next i

        Case 3 ' Three-dimensional array
            ReDim copiedArray(LBound(sourceArray, 1) To UBound(sourceArray, 1), _
                              LBound(sourceArray, 2) To UBound(sourceArray, 2), _
                              LBound(sourceArray, 3) To UBound(sourceArray, 3))
            For i = LBound(sourceArray, 1) To UBound(sourceArray, 1)
                For j = LBound(sourceArray, 2) To UBound(sourceArray, 2)
                    For k = LBound(sourceArray, 3) To UBound(sourceArray, 3)
                        copiedArray(i, j, k) = sourceArray(i, j, k)
                    Next k
                Next j
            Next i

        Case Else
            Err.Raise Number:=vbObjectError + 514, _
                      Description:="Array with more than three dimensions not supported."
    End Select

    ' Return the copied array
    DeepCopyArray = copiedArray
End Function

Function GetStudyNumber() As String
    ' Check if cell A2 is not empty
    If Len(ActiveSheet.Range("A2").Value) > 0 Then
        ' Retrieve and return the first 5 characters of the value in cell A2
        GetStudyNumber = Left(ActiveSheet.Range("A2").Value, 5)
    Else
        ' Return an empty string if A2 is empty
        GetStudyNumber = ""
    End If
End Function

Public Sub QuickRepCleanup()

    'Remove columns from Quick Conmed and AE Reports specified in the array
    
    Dim WS1 As Worksheet
    Dim QuickColName As Variant
    Dim col As Variant
    
    Set WS1 = ActiveSheet

    WS1.Activate
    
    QuickColName = Array("Form Type", "Filled Date", "FK_FILLEDFORM", "FK_FORM", "FORM_COMPLETED", "TYPE", "FK_STUDY", "FK_PER", "ID", "FK_PATPROT", "CREATOR", "EVENT_SCHEDULE")
    
    For Each col In QuickColName
        Call RemoveColumn(WS1, CStr(col))
    Next col
        
End Sub

Public Function HeaderWNoParenthesis(header As String) As String
    Dim pos As Long
    pos = InStrRev(header, "(")
    If pos > 0 Then
        HeaderWNoParenthesis = Trim(Left(header, pos - 1))
    Else
        HeaderWNoParenthesis = header
    End If
End Function

Public Function ExtractInnerString(text As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(1, text, "(", vbTextCompare)
    endPos = InStr(1, text, ")", vbTextCompare)
    
    If startPos > 0 And endPos > startPos Then
        ExtractInnerString = Mid(text, startPos + 1, endPos - startPos - 1)
    Else
        ExtractInnerString = "" ' Return an empty string if not found
    End If
End Function
Sub UpdateSheetFontToCalibri(ws)
    ' Update the entire sheet to Calibri font
    With ws.Cells
        .Font.Name = "Calibri"
    End With
End Sub
