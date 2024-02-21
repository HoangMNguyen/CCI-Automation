VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormattingUI 
   Caption         =   "UserForm1"
   ClientHeight    =   1425
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5580
   OleObjectBlob   =   "FormattingUI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormattingUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Sub UserForm_Initialize()
ComboBox1.Clear
ComboBox1.AddItem "02916"
ComboBox1.AddItem "03818"
ComboBox1.AddItem "12418"
ComboBox1.AddItem "14217"
ComboBox1.AddItem "32816"
ComboBox1.AddItem "35418"
ComboBox1.AddItem "46417"
ComboBox1.AddItem "15CT055"
ComboBox1.AddItem "18CT014"
ComboBox1.AddItem "19CT023"

ComboBox2.Clear
ComboBox2.AddItem "02916-Prior Antineoplastic (ANP) Therapy V2"
ComboBox2.AddItem "02916-Prior Oncology Therapy (run report in linear format)"
ComboBox2.AddItem "02916-Infusion Vital Signs V2-Cohort 5"
ComboBox2.AddItem "02916-Infusion Vital Signs V2-Cohort 6"
ComboBox2.AddItem "02916-Medical History V2"
ComboBox2.AddItem "02916-RECIST V2-Cohort 6"
ComboBox2.AddItem "03818-LTFU Antineoplastic (ANP) Therapy"
ComboBox2.AddItem "03818-Prior Oncology Therapy"
ComboBox2.AddItem "03818-Medical History"
ComboBox2.AddItem "12418-Prior Oncology Therapy"
ComboBox2.AddItem "12418-Medical History V2"
ComboBox2.AddItem "14217-Prior Antineoplastic (ANP) Therapy V2"
ComboBox2.AddItem "14217-Medical History V2"
ComboBox2.AddItem "32816-Prior Oncology Therapy(run report in linear format)"
ComboBox2.AddItem "35418-Bridging Therapy"
ComboBox2.AddItem "35418-Infusion Vital Signs"
ComboBox2.AddItem "35418-Medical History"
ComboBox2.AddItem "35418-Prior Antineoplastic (ANP) Therapy"
ComboBox2.AddItem "35418-Primary Antineoplastic (ANP) Therapy"
ComboBox2.AddItem "35418-LTFU Antineoplastic (ANP) Therapy"
ComboBox2.AddItem "35418-Prior Transplant V2"
ComboBox2.AddItem "35418-Transplant V2"
ComboBox2.AddItem "35418-Transfusion"
ComboBox2.AddItem "46417-Prior Oncology Therapy"
ComboBox2.AddItem "46417-Medical History (run report in linear format)"
ComboBox2.AddItem "15CT055-Prior Antineoplastic (ANP) Therapy V3"
ComboBox2.AddItem "18CT014-Prior ANP Therapy"
ComboBox2.AddItem "18CT014-Primary ANP Therapy"
ComboBox2.AddItem "18CT014-Medical History"
ComboBox2.AddItem "19CT023-Prior ANP Therapy"
ComboBox2.AddItem "19CT023-Medical History"


FormattingUI.CommandButton1.Caption = "Format CRF"
FormattingUI.Caption = "Select the CRF"
End Sub
Sub CommandButton1_Click()

    If ComboBox1.Value = "02916" Then
        If ComboBox2.Value = "02916-Prior Antineoplastic (ANP) Therapy V2" Then  'Prior Antineoplastic (ANP) Therapy V2
            Call S02916.PRIOR_ANP
        ElseIf ComboBox2.Value = "02916-Prior Oncology Therapy (run report in linear format)" Then 'must run the Prior Oncology Therapy report in linear format then export to excel
            Call S02916.PRIOR_ONC
        ElseIf ComboBox2.Value = "02916-Infusion Vital Signs V2-Cohort 6" Then
            Call S02916.INFUSIONVITALS_COHORT6
        ElseIf ComboBox2.Value = "02916-Infusion Vital Signs V2-Cohort 5" Then
            Call S02916.INFUSIONVITALS_COHORT5
        ElseIf ComboBox2.Value = "02916-Medical History V2" Then
            Call S02916.MEDHXV2
        ElseIf ComboBox2.Value = "02916-RECIST V2-Cohort 6" Then
            Call S02916.RECISTV2_COHORT6
        Else
            MsgBox ("Doesn't have this CRF for this study yet. Ask Ming :))")
        End If

    ElseIf ComboBox1.Value = "03818" Then
        If ComboBox2.Value = "03818-LTFU Antineoplastic (ANP) Therapy" Then
            Call S03818.LTFU_ANP
        ElseIf ComboBox2.Value = "03818-Prior Oncology Therapy" Then
            Call S03818.PRIOR_ONC
        ElseIf ComboBox2.Value = "03818-Medical History" Then
            Call S03818.MEDHX
        Else
            MsgBox ("Doesn't have this CRF for this study yet. Ask Ming :))")
        End If
        
    ElseIf ComboBox1.Value = "12418" Then
        If ComboBox2.Value = "12418-Prior Oncology Therapy" Then
            Call S12418.PRIOR_ONC
        ElseIf ComboBox2.Value = "12418-Medical History V2" Then
            Call S12418.MEDHX
        Else
            MsgBox ("Doesn't have this CRF for this study yet. Ask Ming :))")
        End If
    ElseIf ComboBox1.Value = "14217" Then
        If ComboBox2.Value = "14217-Prior Antineoplastic (ANP) Therapy V2" Then
            Call S14217.PRIOR_ANP
        ElseIf ComboBox2.Value = "14217-Medical History V2" Then
            Call S14217.MEDHX
        Else
            MsgBox ("Doesn't have this CRF for this study yet. Ask Ming :))")
        End If
    

    ElseIf ComboBox1.Value = "32816" Then
        If ComboBox2.Value = "32816-Prior Oncology Therapy(run report in linear format)" Then
            Call S32816.PRIOR_ONC
        Else
            MsgBox ("Doesn't have this CRF for this study yet. Ask Ming :))")
        End If

    ElseIf ComboBox1.Value = "19CT023" Then
        If ComboBox2.Value = "19CT023-Prior ANP Therapy" Then
            Call S19CT023.PRIORANP
        ElseIf ComboBox2.Value = "19CT023-Medical History" Then
            Call S19CT023.MEDHX
        Else
            MsgBox ("Doesn't have this CRF for this study yet. Ask Ming :))")
        End If


    ElseIf ComboBox1.Value = "35418" Then
        If ComboBox2.Value = "35418-Bridging Therapy" Then
            Call S35418.BRIDGING_THERAPY
        ElseIf ComboBox2.Value = "35418-LTFU Antineoplastic (ANP) Therapy" Then
            Call S35418.LTFU_ANP
        ElseIf ComboBox2.Value = "35418-Prior Antineoplastic (ANP) Therapy" Then
            Call S35418.PRIOR_ANP
        ElseIf ComboBox2.Value = "35418-Primary Antineoplastic (ANP) Therapy" Then
            Call S35418.PRIMARY_ANP
        ElseIf ComboBox2.Value = "35418-Medical History" Then
            Call S35418.MEDHX
        ElseIf ComboBox2.Value = "35418-Prior Transplant V2" Then
            Call S35418.PRIOR_TRANSPLANT_V2
        ElseIf ComboBox2.Value = "35418-Transplant V2" Then
            Call S35418.TRANSPLANT_V2
        ElseIf ComboBox2.Value = "35418-Transfusion" Then
            Call S35418.TRANSFUSION
        ElseIf ComboBox2.Value = "35418-Infusion Vital Signs" Then
            Call S35418.INFUSIONVITALS
        Else
            MsgBox ("Doesn't have this CRF for this study yet. Ask Ming :))")
        End If
    
    ElseIf ComboBox1.Value = "46417" Then
        If ComboBox2.Value = "46417-Prior Oncology Therapy" Then
            Call S46417.PRIOR_ONC
        ElseIf ComboBox2.Value = "46417-Medical History (run report in linear format)" Then
            Call S46417.MEDHX_L
        Else
            MsgBox ("Doesn't have this CRF for this study yet. Ask Ming :))")
        End If
    ElseIf ComboBox1.Value = "15CT055" Then
        If ComboBox2.Value = "15CT055-Prior Antineoplastic (ANP) Therapy V3" Then
            Call S15CT055.PRIOR_ANP
        Else
            MsgBox ("Doesn't have this CRF for this study yet. Ask Ming :))")
        End If
    ElseIf ComboBox1.Value = "18CT014" Then
        If ComboBox2.Value = "18CT014-Prior ANP Therapy" Then
            Call S18CT014.PRIOR_ANP
        ElseIf ComboBox2.Value = "18CT014-Primary ANP Therapy" Then
            Call S18CT014.PRIMARY_ANP
        ElseIf ComboBox2.Value = "18CT014-Medical History" Then
            Call S18CT014.MEDHX
        Else
            MsgBox ("Doesn't have this CRF for this study yet. Ask Hoang :))")
        End If
    ElseIf ComboBox1.Value = "01817" Then
        If ComboBox2.Value = "01817-Concomitant Medication V1(retired)" Then 'retired, no longer in use
            Call S01817.ConMedLog
        Else
            MsgBox ("Doesn't have this CRF for this study yet. Ask Hoang :))")
        End If
    Else
        MsgBox ("Doesn't have this study.Ask Ming :))")
    
    End If
    Unload FormattingUI
End Sub

