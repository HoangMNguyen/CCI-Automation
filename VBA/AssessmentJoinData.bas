Attribute VB_Name = "AssessmentJoinData"
Option Explicit

Public Sub ProcessAssessmentJoinData()
    Dim sourceWb As Workbook
    Dim outputWb As Workbook
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim questionMap As Object
    Dim groupMap As Object
    Dim questionOrder As Collection
    Dim groupOrder As Collection
    Dim sourceData As Variant
    Dim outputData As Variant
    Dim groupRow As Variant
    Dim groupKey As String
    Dim questionText As String
    Dim subjectCol As Long
    Dim formSeqCol As Long
    Dim statusCol As Long
    Dim questionAnswerCol As Long
    Dim questionTextCol As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim totalColumns As Long
    Dim completedRowCount As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim questionIndex As Long
    Dim finalOutputCol As Long
    Dim saveFolder As String
    Dim saveFilePath As String
    Dim baseFilePath As String

    On Error GoTo CleanFail
    Application.ScreenUpdating = False

    Set sourceWb = ActiveWorkbook
    Set wsSource = ActiveSheet

    wsSource.Activate
    RemoveFilter

    RemoveColumn wsSource, "ASM.GUID"
    RemoveColumn wsSource, "ASMR.AGUID"

    subjectCol = GetColumnNumber(wsSource, "Subject")
    formSeqCol = GetColumnNumber(wsSource, "Form Sequence Number")
    statusCol = GetColumnNumber(wsSource, "ASM.STATUS")
    questionAnswerCol = GetColumnNumber(wsSource, "ASMR.QUESANS")
    questionTextCol = GetColumnNumber(wsSource, "ASMR.QUESTEXT")

    If subjectCol = 0 Or formSeqCol = 0 Or statusCol = 0 _
        Or questionAnswerCol = 0 Or questionTextCol = 0 Then
        MsgBox "One or more required columns are missing from the active sheet.", vbExclamation
        GoTo CleanExit
    End If

    lastRow = FindLastRowA(wsSource)
    If lastRow < 2 Then
        MsgBox "The active sheet does not contain any data rows.", vbExclamation
        GoTo CleanExit
    End If

    lastCol = FindLastColumn(wsSource)
    sourceData = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastCol)).Value2

    Set questionMap = CreateObject("Scripting.Dictionary")
    Set questionOrder = New Collection

    For rowIndex = 2 To UBound(sourceData, 1)
        If StrComp(Trim$(CStr(sourceData(rowIndex, statusCol))), "Completed", vbTextCompare) = 0 Then
            completedRowCount = completedRowCount + 1

            questionText = Trim$(CStr(sourceData(rowIndex, questionTextCol)))
            If Len(questionText) > 0 Then
                If Not questionMap.Exists(questionText) Then
                    questionMap.Add questionText, questionOrder.Count + 1
                    questionOrder.Add questionText
                End If
            End If
        End If
    Next rowIndex

    If completedRowCount = 0 Then
        MsgBox "No rows with ASM.STATUS = Completed were found.", vbInformation
        GoTo CleanExit
    End If

    totalColumns = 2 + questionOrder.Count

    Set groupMap = CreateObject("Scripting.Dictionary")
    Set groupOrder = New Collection

    For rowIndex = 2 To UBound(sourceData, 1)
        If StrComp(Trim$(CStr(sourceData(rowIndex, statusCol))), "Completed", vbTextCompare) = 0 Then
            groupKey = BuildAssessmentJoinKey( _
                sourceData(rowIndex, subjectCol), _
                sourceData(rowIndex, formSeqCol), _
                vbNullString)

            If Not groupMap.Exists(groupKey) Then
                ReDim groupRow(1 To totalColumns)
                groupRow(1) = sourceData(rowIndex, subjectCol)
                groupRow(2) = sourceData(rowIndex, formSeqCol)
                groupMap.Add groupKey, groupRow
                groupOrder.Add groupKey
            End If

            groupRow = groupMap(groupKey)
            questionText = Trim$(CStr(sourceData(rowIndex, questionTextCol)))

            If Len(questionText) > 0 Then
                questionIndex = CLng(questionMap(questionText))
                ' Populate each question column with the matching answer value.
                groupRow(2 + questionIndex) = MergeAssessmentAnswer( _
                    groupRow(2 + questionIndex), _
                    sourceData(rowIndex, questionAnswerCol))
            End If

            groupMap(groupKey) = groupRow
        End If
    Next rowIndex

    Set outputWb = Workbooks.Add(xlWBATWorksheet)
    Set wsOutput = outputWb.Worksheets(1)
    wsOutput.Name = GetAvailableAssessmentSheetName(outputWb, "Assessment Processed")

    ReDim outputData(1 To groupOrder.Count + 1, 1 To totalColumns)

    outputData(1, 1) = "Subject"
    outputData(1, 2) = "Form Sequence Number"

    For colIndex = 1 To questionOrder.Count
        outputData(1, 2 + colIndex) = CStr(questionOrder(colIndex))
    Next colIndex

    For rowIndex = 1 To groupOrder.Count
        groupRow = groupMap(CStr(groupOrder(rowIndex)))
        For colIndex = 1 To totalColumns
            outputData(rowIndex + 1, colIndex) = groupRow(colIndex)
        Next colIndex
    Next rowIndex

    wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(UBound(outputData, 1), totalColumns)).Value = outputData

    SortAssessmentOutput wsOutput, UBound(outputData, 1), totalColumns
    finalOutputCol = FindLastColumn(wsOutput)

    wsOutput.Activate
    wsOutput.Range("A1").Select
    FormatTable
    wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(UBound(outputData, 1), finalOutputCol)).AutoFilter
    wsOutput.Rows(1).WrapText = True

    For colIndex = 1 To questionOrder.Count
        SetWidthWrapColumn wsOutput, CStr(questionOrder(colIndex)), 28
    Next colIndex

    wsOutput.Range("A1").Select

    saveFolder = sourceWb.Path
    If Len(saveFolder) = 0 Then
        saveFolder = Application.DefaultFilePath
    End If

    If Len(saveFolder) > 0 Then
        baseFilePath = saveFolder & Application.PathSeparator & "Assessment final.xlsx"
    Else
        baseFilePath = "Assessment final.xlsx"
    End If

    saveFilePath = GetAvailableSavePath(baseFilePath)
    Application.DisplayAlerts = False
    outputWb.SaveAs fileName:=saveFilePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    MsgBox completedRowCount & " completed rows were combined into " & groupOrder.Count & _
        " rows across " & questionOrder.Count & " assessment questions." & vbNewLine & _
        "Saved as: " & saveFilePath, vbInformation

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Assessment join processing failed: " & Err.Description, vbCritical
End Sub

Private Function BuildAssessmentJoinKey(ByVal subjectValue As Variant, _
                                        ByVal formSequenceValue As Variant, _
                                        ByVal createdValue As Variant) As String
    BuildAssessmentJoinKey = CStr(subjectValue) & Chr$(30) & _
                             CStr(formSequenceValue) & Chr$(30) & _
                             CStr(createdValue)
End Function

Private Function MergeAssessmentAnswer(ByVal existingValue As Variant, ByVal newValue As Variant) As String
    Dim existingText As String
    Dim newText As String

    existingText = Trim$(CStr(existingValue))
    newText = Trim$(CStr(newValue))

    If Len(newText) = 0 Then
        MergeAssessmentAnswer = existingText
    ElseIf Len(existingText) = 0 Then
        MergeAssessmentAnswer = newText
    ElseIf ContainsDelimitedValue(existingText, newText) Then
        MergeAssessmentAnswer = existingText
    Else
        MergeAssessmentAnswer = existingText & " | " & newText
    End If
End Function

Private Function ContainsDelimitedValue(ByVal existingText As String, ByVal candidateText As String) As Boolean
    Dim parts As Variant
    Dim part As Variant

    parts = Split(existingText, " | ")
    For Each part In parts
        If StrComp(Trim$(CStr(part)), candidateText, vbTextCompare) = 0 Then
            ContainsDelimitedValue = True
            Exit Function
        End If
    Next part
End Function

Private Function GetAvailableAssessmentSheetName(ByVal targetWb As Workbook, ByVal baseName As String) As String
    Dim sheetName As String
    Dim counter As Long

    sheetName = Left$(baseName, 31)
    counter = 2

    Do While WorksheetExists(targetWb, sheetName)
        sheetName = Left$(baseName, 31 - Len(CStr(counter)) - 1) & " " & CStr(counter)
        counter = counter + 1
    Loop

    GetAvailableAssessmentSheetName = sheetName
End Function

Private Function WorksheetExists(ByVal targetWb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = targetWb.Worksheets(sheetName)
    On Error GoTo 0

    WorksheetExists = Not ws Is Nothing
End Function

Private Function GetAvailableSavePath(ByVal preferredPath As String) As String
    Dim folderPath As String
    Dim fileNameOnly As String
    Dim dotPos As Long
    Dim fileBase As String
    Dim fileExt As String

    If Dir(preferredPath) = "" Then
        GetAvailableSavePath = preferredPath
        Exit Function
    End If

    folderPath = Left$(preferredPath, InStrRev(preferredPath, Application.PathSeparator))
    fileNameOnly = Mid$(preferredPath, InStrRev(preferredPath, Application.PathSeparator) + 1)
    dotPos = InStrRev(fileNameOnly, ".")

    If dotPos > 0 Then
        fileBase = Left$(fileNameOnly, dotPos - 1)
        fileExt = Mid$(fileNameOnly, dotPos)
    Else
        fileBase = fileNameOnly
        fileExt = ""
    End If

    GetAvailableSavePath = folderPath & fileBase & " " & Now2Date(Now) & "_" & Format(Now, "hhmmss") & fileExt
End Function

Private Sub SortAssessmentOutput(ByVal ws As Worksheet, ByVal lastDataRow As Long, ByVal lastDataCol As Long)
    Dim sortRange As Range

    If lastDataRow < 3 Then Exit Sub

    Set sortRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastDataRow, lastDataCol))

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(2, 1), ws.Cells(lastDataRow, 1)), Order:=xlAscending
        .SortFields.Add Key:=ws.Range(ws.Cells(2, 2), ws.Cells(lastDataRow, 2)), Order:=xlAscending
        .SetRange sortRange
        .Header = xlYes
        .Apply
    End With
End Sub
