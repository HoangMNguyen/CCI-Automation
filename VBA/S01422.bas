Attribute VB_Name = "S01422"
Sub QFSR(WS1 As Worksheet, FilterCriteria1 As String, Optional FilterCriteria2 As String = "")
    
    Dim lastRow As Long
    Dim i As Long
    Dim ws As Worksheet
    Dim dashIndex As Integer
    Dim StudyNum As String
    Dim WSFSO As Worksheet
    lastRow = FindLastRowA(WS1)
    
    Set WSFSO = Sheets.Add(Before:=WS1)
    WSFSO.Name = "Form Status Overview"
    WSFSO.Range("A1").Value = "All Cohorts Form Status"
    Call FormStatusOverview(WS1, WSFSO, 1, lastRow)
    WSFSO.Activate
    Call FormatTable
    
    Dim uniqueValues As Variant
    uniqueValues = ExtractUniqueValues(WS1, 2)
    
    '
    For i = LBound(uniqueValues) To UBound(uniqueValues)   'i starts from 1
        Debug.Print i
        Debug.Print uniqueValues(i, 1)
        Set ws = Sheets.Add(After:=Sheets(i + 1))
        'find where the first dash is
        dashIndex = InStrRev(uniqueValues(i, 1), "-")
        If dashIndex > 0 Then
            StudyNum = Right(uniqueValues(i, 1), Len(uniqueValues(i, 1)) - dashIndex)
        Else
            StudyNum = uniqueValues(i, 1)
        End If
        
        ws.Name = StudyNum & " Form Status"

        'Copy study data to another tab
        WS1.Range("A1").AutoFilter Field:=2, Criteria1:=uniqueValues(i, 1) & "*"
        WS1.Range("A1:S" & lastRow).SpecialCells(xlCellTypeVisible).Copy
        ws.Paste
        Call Main.OutFormat
        
        WSFSO.Range("A" & (1 + i * 7)).Value = StudyNum & " Form Status"
        Call FormStatusOverview(ws, WSFSO, 1 + i * 7, lastRow)
        WSFSO.Activate
        WSFSO.Range("A" & (1 + i * 7)).Select
        Call FormatTable
        If FilterCriteria2 = "" Then
            ws.Range("A1").AutoFilter Field:=6, Criteria1:=FilterCriteria1
        Else
            ws.Range("A1").AutoFilter Field:=6, Criteria1:=FilterCriteria1, Operator:=xlOr, Criteria2:=FilterCriteria2
        End If
    Next i
    
End Sub

Sub QQSR(WS1 As Worksheet)
    
    Dim lastRow As Long
    Dim i As Long
    Dim ws As Worksheet
    Dim dashIndex As Integer
    Dim StudyNum As String
    Dim WSQRO As Worksheet
    lastRow = FindLastRowA(WS1)
    
    Set WSQRO = Sheets.Add(Before:=WS1)
    WSQRO.Name = "Query Report Overview"
    WSQRO.Range("A1").Value = "All Cohorts Query Status"
    Call QueryReportOverview(WS1, WSQRO, 1)
    WSQRO.Activate
    Call FormatTable
    
    Dim uniqueValues As Variant
    uniqueValues = ExtractUniqueValues(WS1, 4)
    
    '
    For i = LBound(uniqueValues) To UBound(uniqueValues)   'i starts from 1
        Debug.Print i
        Debug.Print uniqueValues(i, 1)
        Set ws = Sheets.Add(After:=Sheets(i + 1))
        'find where the first dash is
        dashIndex = InStrRev(uniqueValues(i, 1), "-")
        If dashIndex > 0 Then
            StudyNum = Right(uniqueValues(i, 1), Len(uniqueValues(i, 1)) - dashIndex)
        Else
            StudyNum = uniqueValues(i, 1)
        End If
        
        ws.Name = StudyNum & " Query Report"

        'Copy study data to another tab
        WS1.Range("A1").AutoFilter Field:=4, Criteria1:=uniqueValues(i, 1) & "*"
        WS1.Range("A1:T" & lastRow).SpecialCells(xlCellTypeVisible).Copy
        ws.Paste
        Call Main.OutFormat
        
        WSQRO.Range("A" & (1 + i * 6)).Value = StudyNum & " Query Status"
        Call QueryReportOverview(ws, WSQRO, 1 + i * 6)
        WSQRO.Activate
        WSQRO.Range("A" & (1 + i * 6)).Select
        Call FormatTable

    Next i
    
End Sub

