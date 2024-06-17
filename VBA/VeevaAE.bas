Attribute VB_Name = "VeevaAE"
Sub FormatVeevaAE() 'Reformat Veeva Core listing for AE report to match the safety team request. This works for AE v1: 15420 study

    Dim WB1 As Workbook
    Dim WS1 As Worksheet
    Dim WS2 As Worksheet
    Dim lastRow As Long
    Dim StudyNum As String
    Dim HeadersIndex As Variant
    
    Application.ScreenUpdating = False
    
    Set WB1 = ActiveWorkbook
    Set WS1 = ActiveSheet
    ActiveSheet.Name = "Veeva AE Data Listing"
    'Count number of rows w/ header
    WS1.Range("A1").Select
    lastRow = ActiveSheet.Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
    


    'Add WS2 as output
    With WB1
        .Sheets.Add(Before:=.Sheets(.Sheets.Count)).Name = "Reformated AE Report"
        Set WS2 = Sheets("Reformated AE Report")
    End With

    StudyNum = Left(WS1.Range("A2").Value, 5)
    If Len(StudyNum) <> 5 And Not IsNumeric(StudyNum) Then
        StudyNum = Left(ActiveWorkbook.Name, 5)
    End If
    Call CopyColumnsWithHeaders(WS1, WS2, GetAEHeaders(StudyNum), 1, 1)
    HeadersIndex = FindHeaderIndexes(WS1, GetAEHeaders(StudyNum))
    
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
    
    WS2.Cells(1, 1).Select
    Call FormatTable
    With ActiveWindow
        .SplitRow = 1
    End With
    
    ActiveWindow.FreezePanes = True
    Application.ScreenUpdating = True
    
    Call WarningCSV

End Sub

Function GetAEHeaders(StudyNum As String) As Variant

    If StudyNum = "15420" Then
        GetAEHeaders = Array("Subject", "AE or SAE? (ig_AE2.AESEV)", "T-cell Attribution (ig_AE1.AEREL)", "T-cell Expectedness (ig_AE1.AETRTINTP)", "Specify Other Attribution (ig_AE1.AERELSPOTH)", "Other Attribution (ig_AE1.AERELOTH)", "Other Expectedness (ig_AE1.AETRTINTPOTH)", "CTCAE Category (ig_AE1.AECAT)", "Toxicity (ig_AE1.AETOX)", "Grade (ig_AE1.AETOXGR)", "Start Date (ig_AE1.AESTDAT)", "Stop Date (ig_AE1.AEENDAT)", "Event Onset (ig_AE1.AEONSET)", "Additional Toxicity Details (ig_AE1.AETOXTERM)", "Event Ongoing (ig_AE1.AEONGO)")
    ElseIf StudyNum = "03821" Then
        GetAEHeaders = Array("Subject", "Cohort Assignment (ig_AE1.CACHASCOD)", "AE or SAE? (ig_AE2.AESEV)", "Investigational Product(s) (ig_AE1.AEIP)", "Attribution to T-cell Therapy (IP1) (ig_AE1.AEREL1)", "T-cell Therapy Expectedness (IP1) (ig_AE1.AETRTINTP1)", "Attribution to VCN-01 (IP2) (ig_AE1.AEREL2)", "VCN-01 Expectedness (IP2) (ig_AE1.AETRTINTP2)", "Other Attribution (ig_AE1.AERELOTH)", "Specify Other Attribution (ig_AE1.AERELSPOTH)", "CTCAE Category (ig_AE1.AECAT)", "Toxicity (ig_AE1.AETOX)", "Derived Toxicity (ig_AE1.AETOXDV)", "Grade (ig_AE1.AETOXGR)", "Start Date (ig_AE1.AESTDAT)", "Stop Date (ig_AE1.AEENDAT)", "Event Onset (ig_AE1.AEONSETCH-1)", "Event Onset (ig_AE1.AEONSETCH12)", "Event Onset (ig_AE1.AEONSETCHNAS)", "Event Ongoing? (ig_AE1.AEONGO)", "Additional Toxicity Details (ig_AE1.AETOXTERM)")
    ElseIf StudyNum = "16321" Then
        GetAEHeaders = Array("Subject", "AE or SAE? (ig_AE2.AESEV)", "T-cell Attribution (ig_AE1.AEREL)", "T-cell Expectedness (ig_AE1.AETRTINTP)", "Specify Other Attribution (ig_AE1.AERELSPOTH)", "Other Attribution (ig_AE1.AERELOTH)", "Other Expectedness (ig_AE1.AETRTINTPOTH)", "CTCAE Category (ig_AE1.AECAT)", "Derived Toxicity (ig_AE1.AETOXDV)", "Toxicity (ig_AE1.AETOX)", "Grade (ig_AE1.AETOXGR)", "Start Date (ig_AE1.AESTDAT)", "Stop Date (ig_AE1.AEENDAT)", "Event Onset (ig_AE1.AEONSET)", "Additional Toxicity Details (ig_AE1.AETOXTERM)", "Event Ongoing (ig_AE1.AEONGO)")
    ElseIf StudyNum = "15122" Then
        GetAEHeaders = Array("Subject", "AE or SAE? (IG_NS_NA_AE2.CL_YS_YH_AESEV_cl_YS_AESAE1)", "T-cell Attribution (IG_NS_NA_AE1.CL_YS_NH_AEREL_cl_YS_TCELLATRIB1)", "T-cell Expectedness (IG_NS_NA_AE1.CL_YS_YH_AETRTINTP_cl_YS_YN1)", "Other Attribution (IG_NS_NA_AE1.CL_YS_NH_RELOTH_cl_YS_OTHATRIB1)", "Specify Other Attribution (IG_NS_NA_AE1.TX_YS_NH_AERELSPOTH)", "CTCAE Category (IG_NS_NA_AE1.CL_YS_NH_AECAT_cl_YS_CTCAECAT2)", "CTCAE Term Available? (IG_NS_NA_AE1.CL_YS_YH_AECATOTH_cl_YS_YN1)", "Derived Toxicity (IG_NS_NA_AE1.DV_YS_YH_AETOXDV)", "Toxicity (IG_NS_NA_AE1.TX_YS_NH_AETOX)", "Grade (IG_NS_NA_AE1.CL_YS_YH_AETOXGR_cl_YS_AEGRADE1)", "Start Date (IG_NS_NA_AE1.DT_YS_NH_AESTDAT)", "Stop Date (IG_NS_NA_AE1.DT_YS_YH_AEENDAT)", "Event Onset (IG_NS_NA_AE1.CL_NS_YH_AEONSET_cl_NS_AEONSET1)", "Event Ongoing (IG_NS_NA_AE1.CL_YS_YH_AEONGO_cl_YS_AEONGO1)", "Additional Toxicity Details (IG_NS_NA_AE1.TX_YS_YH_AETOXTERM)")
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
