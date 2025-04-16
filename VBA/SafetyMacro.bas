Attribute VB_Name = "SafetyMacro"
Sub FormatNonAECRF()
'Last updated: 04/14/2025 By Nina Sizova

    Dim xcell As Object
    Dim SelectCells As Range
    Dim lastRow As Long
    Dim LastCol As Long
    Dim LastColumn As String
    Dim WS1 As Worksheet
    Dim WS2 As Worksheet
    
    'Disable Screen Update
    Application.ScreenUpdating = False
    
    'Activate workbook
    ActiveWorkbook.Activate
    Set WS1 = Sheets(1)
    WS1.Activate
    
    WS1.Copy Before:=Sheets(1)
    Set WS2 = Sheets(1)
    WS2.Name = "DATA"
    
    'Autofit from A to last column
    WS2.UsedRange.Columns.AutoFit
    Call RemoveColumn(WS2, "Study")
    Call RemoveColumn(WS2, "Study Country")
    Call RemoveColumn(WS2, "Study Site")
    
    ActiveSheet.Range("A1").Select
    Call FormatTable
    
    WS2.Activate
    WS2.Range("A1").Select
    'Find last row
    lastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
    'Find last column then split it to find the letter
    LastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    LastColumn = GetColumnLetter(LastCol)

    
    'Hidden Range (Event Group Name to Data Listing As Of)
    Dim EGNColumn As String
    Dim DLAOColumn As String
    Dim FNColumn As String
    Dim FSColumn As String
    EGNColumn = FindColumn(WS2, "Event Group Name")
    DLAOColumn = FindColumn(WS2, "Data Listing As Of")
    FNColumn = FindColumn(WS2, "Form Name")
    FSColumn = FindColumn(WS2, "Form Status")
    Columns(EGNColumn & ":" & FNColumn).Hidden = True
    Columns(FSColumn & ":" & DLAOColumn).Hidden = True
    If Not IsEmpty(Range("A2")) Then
        WS2.Range("A2:" & LastColumn & lastRow).Select
        Selection.VerticalAlignment = xlTop
        Selection.WrapText = True
    End If
    
    'Sort by Form Sequence Number and Subject
    Call sortAsc("K1")
    Call sortAsc("A1")
    
    'Freeze top row
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    If Not IsEmpty(Range("A2")) Then
       Range("A2").Select
       ActiveWindow.FreezePanes = True
       'Sort
       Range("A1:" & LastColumn & lastRow).Sort Key1:=Range("A1:" & LastColumn & lastRow), Order1:=xlAscending, header:=xlYes
    End If
    'Filter
    WS2.UsedRange.Select
    Selection.AutoFilter

    'Enable Screen Update
    Application.ScreenUpdating = True
    'Check CSV
    Call WarningCSV
    
End Sub
