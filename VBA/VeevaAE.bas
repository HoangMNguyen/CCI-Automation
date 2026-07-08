Attribute VB_Name = "VeevaAE"
Sub FormatVeevaAE() 'Reformat Veeva Core listing for AE report to match the safety team request. This works for AE v1: 15420 study
    'Last update: 6/11/2025 by Hoang Nguyen
    Dim WB1 As Workbook
    Dim WS1 As Worksheet
    Dim WS2 As Worksheet
    Dim lastRow As Long
    Dim StudyNum As String
    Dim HeadersIndex As Variant
    
    Application.ScreenUpdating = False
    
    Set WB1 = ActiveWorkbook
    Set WS1 = ActiveSheet
    'Count number of rows w/ header
    lastRow = FindLastRowA(WS1)
    
    StudyNum = Left(WS1.Range("A2").Value, 5)
    If Len(StudyNum) <> 5 And Not IsNumeric(StudyNum) Then
        StudyNum = Left(ActiveWorkbook.Name, 5)
    End If
    ' TODO:
    'Determine which set of headers to use based on the "Form Name" column
    Dim FormNameCol As String
    Dim FormNameValue As String
    Dim SelectedHeaders As Variant
    Dim ReportSheetName As String
    Dim ListingSheetName As String
    FormNameCol = FindColumn(WS1, "Form Name")
    If FormNameCol <> "Not Found" Then
        FormNameValue = CStr(WS1.Range(FormNameCol & "2").Value)
    End If
    'If the form name contains PDAE use the PDAE headers, otherwise (AE) use the AE headers
    If InStr(1, FormNameValue, "PDAE", vbTextCompare) > 0 Then
        SelectedHeaders = GetPDAEHeaders(StudyNum)
        ReportSheetName = "Reformated PDAE Report"
        ListingSheetName = "Veeva PDAE Data Listing"
    Else
        SelectedHeaders = GetAEHeaders(StudyNum)
        ReportSheetName = "Reformated AE Report"
        ListingSheetName = "Veeva AE Data Listing"
    End If
    'Rename the source listing sheet based on the form type
    WS1.Name = ListingSheetName

    'Add WS2 as output
    With WB1
        .Sheets.Add(Before:=.Sheets(.Sheets.count)).Name = ReportSheetName
        Set WS2 = Sheets(ReportSheetName)
    End With

    Call CopyColumnsWithHeaders(WS1, WS2, SelectedHeaders, 1, 1)
    HeadersIndex = FindHeaderIndexes(WS1, SelectedHeaders)
    Dim index As Long
    For index = 1 To UBound(HeadersIndex) + 1
        'Split header name to remove part after the last "("
        WS2.Cells(1, index).Value = HeaderWNoParenthesis(WS2.Cells(1, index).Value)
    Next index
    
    'Copy the rest of the form
    Dim LastColumn As Long
    Dim LastColumnWS2 As Long
    LastColumn = FindLastColumn(WS1)
    LastColumnWS2 = FindLastColumn(WS2)
    
    Dim i As Long
    Dim j As Long
    j = 0
    For i = 1 To LastColumn
        If Not IsNumberInArray(i, HeadersIndex) Then
            Call CopyColumnsWithIndex(WS1, WS2, i, i + LastColumnWS2 - j)
        Else
            j = j + 1
        End If
    Next i
    'Calculate the Duration column
    'Find the start date column
    Dim StartDate As String
    StartDate = FindColumn(WS2, "Start Date")
    Dim StopDate As String
    StopDate = FindColumn(WS2, "Stop Date")
    Set startDateRange = WS2.Range(StartDate & "2:" & StartDate & lastRow)
    Set stopDateRange = WS2.Range(StopDate & "2:" & StopDate & lastRow)
    Dim DurationColumn As Long
    DurationColumn = stopDateRange.column + 1
    'Insert empty column
    WS2.Columns(DurationColumn).Insert Shift = xlToRight
    ' Set the header for the duration column
    WS2.Cells(1, DurationColumn).Value = "Duration"
    'Check if there is no data
    If lastRow > 1 Then
        Set durationRange = WS2.Cells(2, DurationColumn).Resize(lastRow - 1, 1)
        ' Read data into arrays
        Dim startDateArray As Variant
        Dim stopDateArray As Variant
        Dim durationArray() As Variant
        startDateArray = startDateRange.Value
        stopDateArray = stopDateRange.Value
        ' Prepare the duration array
        If IsArray(startDateArray) Then
            If UBound(startDateArray, 1) = 1 And UBound(startDateArray, 2) = 1 Then
                ' Handle special case where there is only one row of data
                ReDim durationArray(1 To 1, 1 To 1)
            Else
                ' General case for arrays with more than one element
                ReDim durationArray(1 To UBound(startDateArray, 1), 1 To 1)
            End If
            
            ' Perform the calculation
            For i = 1 To UBound(startDateArray, 1)
                If IsDate(startDateArray(i, 1)) And IsDate(stopDateArray(i, 1)) Then
                    durationArray(i, 1) = DateDiff("d", startDateArray(i, 1), stopDateArray(i, 1)) + 1
                Else
                    durationArray(i, 1) = "" ' Empty field
                End If
            Next i
        Else
            ' Handle single value case
            ReDim durationArray(1 To 1, 1 To 1)
            If IsDate(startDateArray) And IsDate(stopDateArray) Then
                durationArray(1, 1) = DateDiff("d", startDateArray, stopDateArray) + 1
            Else
                durationArray(1, 1) = "" ' Empty field
            End If
        End If
    ' Write the results back to the worksheet
    durationRange.Value = durationArray
    ' Set the number format of the duration column to number
    durationRange.EntireColumn.NumberFormat = "0"
    End If
    ' Set the color of the new column to HEX #6666FF
    WS2.Cells(2, DurationColumn).EntireColumn.Font.Color = RGB(102, 102, 255)
    
    
    
    'for study 15420 Derived Toxicity rule only
    If StudyNum = "15420" Then
        Dim innerString As String
        'Find column of Toxicity
        Dim Toxicity As String
        Toxicity = FindColumn(WS2, "Toxicity")
        Set toxicityRange = WS2.Range(Toxicity & "2:" & Toxicity & lastRow)
        ' Read data into arrays
        toxicityArray = toxicityRange.Value
        Dim derivedToxicityColumn As Long
        derivedToxicityColumn = toxicityRange.column
        WS2.Columns(derivedToxicityColumn).Insert Shift = xlToRight
        ' Define the derived toxicity range after the insert
        Set derivedToxicityRange = WS2.Range(WS2.Cells(2, derivedToxicityColumn), WS2.Cells(lastRow, derivedToxicityColumn)) ' Define the derived toxicity range
        ReDim derivedToxicityArray(1 To UBound(toxicityArray, 1), 1 To 1)
        ' Perform the calculation
        For i = 1 To UBound(toxicityArray, 1)
            Toxicity = toxicityArray(i, 1)
            derivedToxicity = Toxicity ' Initialize derived toxicity with original toxicity value
            ' Check if "Other" word is present and apply rules
            If InStr(1, Toxicity, "Other", vbTextCompare) > 0 Then
                ' Extract the text within parentheses
                innerString = ExtractInnerString(Toxicity)
                ' Apply the rules
                If Left(UCase(innerString), 4) = "CAR " Then
                    derivedToxicity = Replace(Toxicity, innerString, "CAR " & LCase(Mid(innerString, 5)))
                Else
                    If Left(UCase(innerString), 5) = "COVID" Then
                        derivedToxicity = Toxicity
                    Else
                        derivedToxicity = Replace(Toxicity, innerString, UCase(Left(innerString, 1)) & LCase(Mid(innerString, 2)))
                    End If
                End If
                derivedToxicity = Replace(derivedToxicity, "hlh", "HLH")
            End If
            ' Store the result in the derived toxicity array
            derivedToxicityArray(i, 1) = derivedToxicity
        Next i
        ' Write the results back to the worksheet
        derivedToxicityRange.Value = derivedToxicityArray
        ' Set the header for the derived toxicity column
        WS2.Cells(1, derivedToxicityColumn).Value = "Derived Toxicity"
        ' Set the font color of the new column to HEX #6666FF
        derivedToxicityRange.EntireColumn.Font.Color = RGB(102, 102, 255)
    End If
    
    Dim CTCAEabbrev As Variant
    CTCAEabbrev = Array("COVID", "GGT ", "INR ", "CD4 ", "CPK ", " I ", " II ", " T ", " QT ", " NOS", "CAR ", "HLH")
    'Find column of Derived Toxicity
    Dim DeTox As String
    DeTox = FindColumn(WS2, "Derived Toxicity")
    Set DeToxRange = WS2.Range(DeTox & "2:" & DeTox & lastRow)
    ' Read data into arrays
    DeToxArray = DeToxRange.Value
    Dim NewDeToxValue As String
    Dim index2 As Long
    If IsArray(DeToxArray) Then
        For i = 1 To UBound(DeToxArray, 1)
            If Not IsEmpty(DeToxArray(i, 1)) Then
                NewDeToxValue = CStr(DeToxArray(i, 1))
                For j = 0 To UBound(CTCAEabbrev, 1)
                    index2 = InStr(1, UCase(NewDeToxValue), CTCAEabbrev(j), vbTextCompare)
                    If index2 > 0 Then
                        DeToxArray(i, 1) = Left(NewDeToxValue, index2 - 1) & CTCAEabbrev(j) & Mid(NewDeToxValue, index2 + Len(CTCAEabbrev(j)))
                    End If
                Next j
            End If
        Next i
    Else
        NewDeToxValue = DeToxArray
        For j = 0 To UBound(CTCAEabbrev, 1)
            index2 = InStr(1, UCase(NewDeToxValue), CTCAEabbrev(j), vbTextCompare)
            If index2 > 0 Then
                DeToxArray(1, 1) = Left(NewDeToxValue, index2 - 1) & CTCAEabbrev(j) & Mid(NewDeToxValue, index2 + Len(CTCAEabbrev(j)))
            End If
        Next j
    End If
            
    ' Write the results back to the worksheet
    DeToxRange.Value = DeToxArray
    DeToxRange.EntireColumn.Font.Color = RGB(102, 102, 255)
    
    
    ' Remove columns whose headers end with "_RAW"
    Dim col As Long
    Dim lastColWS2AfterProcessing As Long
    lastColWS2AfterProcessing = FindLastColumn(WS2)
    For col = lastColWS2AfterProcessing To 1 Step -1
        If Right(Trim(WS2.Cells(1, col).Value), 4) = "_RAW" Then
            WS2.Columns(col).Delete
        End If
    Next col
    
    'Formatting
    Call UpdateSheetFontToCalibri(WS2)
    WS2.Cells(1, 1).Select
    Call FormatTable
    With ActiveWindow
        .SplitRow = 1
    End With
    Selection.AutoFilter
    
    ActiveWindow.FreezePanes = True
    Application.ScreenUpdating = True
    
    Call WarningCSV

End Sub

Function GetAEHeaders(StudyNum As String) As Variant

    If StudyNum = "15420" Then
        GetAEHeaders = Array("Subject", _
                            "AE or SAE? (ig_AE2.AESEV)", _
                            "T-cell Attribution (ig_AE1.AEREL)", _
                            "T-cell Expectedness (ig_AE1.AETRTINTP)", _
                            "Specify Other Attribution (ig_AE1.AERELSPOTH)", _
                            "Other Attribution (ig_AE1.AERELOTH)", _
                            "Other Expectedness (ig_AE1.AETRTINTPOTH)", _
                            "CTCAE Category (ig_AE1.AECAT)", _
                            "Toxicity (ig_AE1.AETOX)", _
                            "Grade (ig_AE1.AETOXGR)", _
                            "Start Date (ig_AE1.AESTDAT)", _
                            "Stop Date (ig_AE1.AEENDAT)", _
                            "Event Onset (ig_AE1.AEONSET)", _
                            "Additional Toxicity Details (ig_AE1.AETOXTERM)", _
                            "Event Ongoing (ig_AE1.AEONGO)")
    ElseIf StudyNum = "03821" Then
        GetAEHeaders = Array("Subject", _
                            "Cohort/Treatment Arm Assignment (ig_AE1.CACHASCOD)", _
                            "AE or SAE? (ig_AE2.AESEV)", _
                            "Investigational Product(s) (ig_AE1.AEIP)", _
                            "Attribution to T-cell Therapy (IP1) (ig_AE1.AEREL1)", _
                            "T-cell Therapy Expectedness (IP1) (ig_AE1.AETRTINTP1)", _
                            "Attribution to VCN-01 (IP2) (ig_AE1.AEREL2)", _
                            "VCN-01 Expectedness (IP2) (ig_AE1.AETRTINTP2)", _
                            "Other Attribution (ig_AE1.AERELOTH)", _
                            "Specify Other Attribution (ig_AE1.AERELSPOTH)", _
                            "CTCAE Category (ig_AE1.AECAT)", _
                            "Toxicity (ig_AE1.AETOX)", _
                            "Derived Toxicity (ig_AE1.AETOXDV)", _
                            "Grade (ig_AE1.AETOXGR)", _
                            "Start Date (ig_AE1.AESTDAT)", _
                            "Stop Date (ig_AE1.AEENDAT)", _
                            "Event Onset (ig_AE1.AEONSETCH-1)", _
                            "Event Onset (ig_AE1.AEONSETCH12)", _
                            "Event Onset (ig_AE1.AEONSETCHNAS)", _
                            "Event Ongoing? (ig_AE1.AEONGO)", _
                            "Additional Toxicity Details (ig_AE1.AETOXTERM)")
    ElseIf StudyNum = "03325" Then
        GetAEHeaders = Array("Subject", _
                            "AE or SAE (IG_NS_NA_AE2.CL_NS_YH_AESEV_cl_NS_AESAE1)", _
                            "T-cell Attribution (IG_NS_NA_AE1.CL_YS_NH_AEREL_cl_NS_TCELLATRIB1)", _
                            "T-cell Expectedness (IG_NS_NA_AE1.CL_NS_YH_AETRTINTP_cl_NS_EU1)", _
                            "Specify Other Attribution (IG_NS_NA_AE1.TX_YS_YH_AERELSPOTH)", _
                            "Other Attribution (IG_NS_NA_AE1.CL_YS_NH_RELOTH_cl_NS_OTHATRIB1)", _
                            "CTCAE Category (IG_NS_NA_AE1.CL_YS_NH_AECAT_cl_NS_CTCAECAT2)", _
                            "Derived Toxicity (IG_NS_NA_AE1.DV_YS_YH_AETOXDV)", _
                            "Toxicity (IG_NS_NA_AE1.TX_YS_NH_AETOX)", _
                            "Grade (IG_NS_NA_AE1.CL_NS_YH_AETOXGR_cl_YS_AEGRADE1)", _
                            "Start Date (IG_NS_NA_AE1.DT_YS_NH_AESTDAT)", _
                            "Stop Date (IG_NS_NA_AE1.DT_YS_YH_AEENDAT)", _
                            "Event Onset (IG_NS_NA_AE1.CL_NS_YH_AEONSET_cl_NS_AEONSET1)", _
                            "Additional Toxicity Details (IG_NS_NA_AE1.TX_NS_YH_AETOXTERM)", _
                            "Event Ongoing (IG_NS_NA_AE1.CL_YS_YH_AEONGO_cl_NS_AEONGO1)")
    ElseIf StudyNum = "10325" Then
        GetAEHeaders = Array("Subject", _
                            "AE or SAE (IG_NS_NA_AE2.CL_NS_YH_AESEV_cl_NS_AESAE1)", _
                            "T-cell Attribution (IG_NS_NA_AE1.CL_NS_NH_AEREL_cl_NS_TCELLATRIB1)", _
                            "T-cell Expectedness (IG_NS_NA_AE1.CL_NS_YH_AETRTINTP_cl_NS_EU1)", _
                            "Specify Other Attribution (IG_NS_NA_AE1.TX_YS_YH_AERELSPOTH)", _
                            "Other Attribution (IG_NS_NA_AE1.CL_YS_NH_RELOTH_cl_NS_OTHATRIB1)", _
                            "CTCAE Category (IG_NS_NA_AE1.CL_YS_NH_AECAT_cl_NS_CTCAECAT2)", _
                            "Derived Toxicity (IG_NS_NA_AE1.DV_YS_YH_AETOXDV)", _
                            "Toxicity (IG_NS_NA_AE1.TX_YS_NH_AETOX)", _
                            "Grade (IG_NS_NA_AE1.CL_NS_YH_AETOXGR_cl_YS_AEGRADE1)", _
                            "Start Date (IG_NS_NA_AE1.DT_YS_NH_AESTDAT)", _
                            "Stop Date (IG_NS_NA_AE1.DT_YS_YH_AEENDAT)", _
                            "Event Onset (IG_NS_NA_AE1.CL_NS_YH_AEONSET_cl_NS_AEONSET1)", _
                            "Additional Toxicity Details (IG_NS_NA_AE1.TX_NS_YH_AETOXTERM)", _
                            "Event Ongoing (IG_NS_NA_AE1.CL_YS_YH_AEONGO_cl_NS_AEONGO1)")
    ElseIf StudyNum = "16321" Then
        GetAEHeaders = Array("Subject", _
                            "AE or SAE? (ig_AE2.AESEV)", _
                            "T-cell Attribution (ig_AE1.AEREL)", _
                            "T-cell Expectedness (ig_AE1.AETRTINTP)", _
                            "Specify Other Attribution (ig_AE1.AERELSPOTH)", _
                            "Other Attribution (ig_AE1.AERELOTH)", _
                            "Other Expectedness (ig_AE1.AETRTINTPOTH)", _
                            "CTCAE Category (ig_AE1.AECAT)", _
                            "Derived Toxicity (ig_AE1.AETOXDV)", _
                            "Toxicity (ig_AE1.AETOX)", _
                            "Grade (ig_AE1.AETOXGR)", _
                            "Start Date (ig_AE1.AESTDAT)", _
                            "Stop Date (ig_AE1.AEENDAT)", _
                            "Event Onset (ig_AE1.AEONSET)", _
                            "Additional Toxicity Details (ig_AE1.AETOXTERM)", _
                            "Event Ongoing (ig_AE1.AEONGO)")
    ElseIf StudyNum = "15122" Then
        GetAEHeaders = Array("Subject", "AE or SAE? (IG_NS_NA_AE2.CL_YS_YH_AESEV_cl_YS_AESAE1)", _
                            "T-cell Attribution (IG_NS_NA_AE1.CL_YS_NH_AEREL_cl_YS_TCELLATRIB1)", _
                            "T-cell Expectedness (IG_NS_NA_AE1.CL_YS_YH_AETRTINTP_cl_YS_YN1)", _
                            "Other Attribution (IG_NS_NA_AE1.CL_YS_NH_RELOTH_cl_YS_OTHATRIB1)", _
                            "Specify Other Attribution (IG_NS_NA_AE1.TX_YS_NH_AERELSPOTH)", _
                            "CTCAE Category (IG_NS_NA_AE1.CL_YS_NH_AECAT_cl_YS_CTCAECAT2)", _
                            "CTCAE Term Available? (IG_NS_NA_AE1.CL_YS_YH_AECATOTH_cl_YS_YN1)", _
                            "Derived Toxicity (IG_NS_NA_AE1.DV_YS_YH_AETOXDV)", _
                            "Toxicity (IG_NS_NA_AE1.TX_YS_NH_AETOX)", _
                            "Grade (IG_NS_NA_AE1.CL_YS_YH_AETOXGR_cl_YS_AEGRADE1)", _
                            "Start Date (IG_NS_NA_AE1.DT_YS_NH_AESTDAT)", _
                            "Stop Date (IG_NS_NA_AE1.DT_YS_YH_AEENDAT)", _
                            "Event Onset (IG_NS_NA_AE1.CL_NS_YH_AEONSET_cl_NS_AEONSET1)", _
                            "Event Ongoing (IG_NS_NA_AE1.CL_YS_YH_AEONGO_cl_YS_AEONGO1)", _
                            "Additional Toxicity Details (IG_NS_NA_AE1.TX_YS_YH_AETOXTERM)")
    ElseIf StudyNum = "12423" Then
        GetAEHeaders = Array("Subject", _
                            "AE or SAE? (IG_NS_NA_AE2.CL_YS_YH_AESEV_cl_NS_AESAE1)", _
                            "T-cell Attribution (IG_NS_NA_AE1.CL_YS_NH_AEREL_cl_NS_TCELLATRIB1)", _
                            "T-cell Expectedness (IG_NS_NA_AE1.CL_YS_YH_AETRTINTP_cl_YS_YN1)", _
                            "Specify Other Attribution (IG_NS_NA_AE1.TX_YS_NH_AERELSPOTH)", _
                            "Other Attribution (IG_NS_NA_AE1.CL_YS_NH_RELOTH_cl_NS_OTHATRIB1)", _
                            "CTCAE Category (IG_NS_NA_AE1.CL_YS_NH_AECAT_cl_NS_CTCAECAT2)", _
                            "Derived Toxicity (IG_NS_NA_AE1.DV_YS_YH_AETOXDV)", _
                            "Toxicity (IG_NS_NA_AE1.TX_YS_NH_AETOX)", _
                            "Grade (IG_NS_NA_AE1.CL_YS_YH_AETOXGR_cl_YS_AEGRADE1)", _
                            "Start Date (IG_NS_NA_AE1.DT_YS_NH_AESTDAT)", _
                            "Stop Date (IG_NS_NA_AE1.DT_YS_YH_AEENDAT)", _
                            "Event Onset (IG_NS_NA_AE1.CL_NS_YH_AEONSET_cl_NS_AEONSET1)", _
                            "Additional Toxicity Details (IG_NS_NA_AE1.TX_YS_YH_AETOXTERM)", _
                            "Event Ongoing (IG_NS_NA_AE1.CL_YS_YH_AEONGO_cl_NS_AEONGO1)")
    ElseIf StudyNum = "11823" Then
        GetAEHeaders = Array("Subject", _
                            "AE or SAE? (IG_NS_NA_AE2.CL_YS_YH_AESEV_cl_NS_AESAE1)", _
                            "T-cell Attribution (IG_NS_NA_AE1.CL_YS_NH_AEREL_cl_NS_TCELLATRIB1)", _
                            "T-cell Expectedness (IG_NS_NA_AE1.CL_NS_YH_AETRTINTP_cl_YS_YN1)", _
                            "Specify Other Attribution (IG_NS_NA_AE1.TX_YS_NH_AERELSPOTH)", _
                            "Other Attribution (IG_NS_NA_AE1.CL_YS_NH_RELOTH_cl_NS_OTHATRIB1)", _
                            "CTCAE Category (IG_NS_NA_AE1.CL_YS_NH_AECAT_cl_NS_CTCAECAT2)", _
                            "Derived Toxicity (IG_NS_NA_AE1.DV_YS_YH_AETOXDV)", _
                            "Toxicity (IG_NS_NA_AE1.TX_YS_NH_AETOX)", _
                            "Grade (IG_NS_NA_AE1.CL_YS_YH_AETOXGR_cl_YS_AEGRADE1)", _
                            "Start Date (IG_NS_NA_AE1.DT_YS_NH_AESTDAT)", _
                            "Stop Date (IG_NS_NA_AE1.DT_YS_YH_AEENDAT)", _
                            "Event Onset (IG_NS_NA_AE1.CL_NS_YH_AEONSET_cl_NS_AEONSET1)", _
                            "Additional Toxicity Details (IG_NS_NA_AE1.TX_YS_YH_AETOXTERM)", _
                            "Event Ongoing (IG_NS_NA_AE1.CL_YS_YH_AEONGO_cl_NS_AEONGO1)")
    End If

End Function

Function GetPDAEHeaders(StudyNum As String) As Variant

    If StudyNum = "15420" Then
        GetPDAEHeaders = Array("Subject", _
                               "AE or SAE? (ig_PDAE2.AESEV)", _
                               "T-cell Attribution (ig_PDAE1.AEREL)", _
                               "T-cell Expectedness (ig_PDAE1.AETRTINTP)", _
                               "Other Attribution (ig_PDAE1.AERELOTH)", _
                               "Specify Other Attribution (ig_PDAE1.AERELSPOTH)", _
                               "Other Expectedness (ig_PDAE1.AETRTINTPOTH)", _
                               "CTCAE Category (ig_PDAE1.AECAT)", _
                               "Toxicity (ig_PDAE1.AETOX)", _
                               "Grade (ig_PDAE1.AETOXGR)", _
                               "Start Date (ig_PDAE1.AESTDAT)", _
                               "Stop Date (ig_PDAE1.AEENDAT)", _
                               "Additional Toxicity Details (ig_PDAE1.AETOXTERM)", _
                               "Event Ongoing (ig_PDAE1.AEONGO)", _
                               "Event Onset (ig_PDAE1.PDAEONSET)")
    End If

End Function
