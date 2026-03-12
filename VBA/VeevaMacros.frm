VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VeevaMacros 
   Caption         =   "Veeva Automation"
   ClientHeight    =   1560
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   5100
   OleObjectBlob   =   "VeevaMacros.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VeevaMacros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Sub UserForm_Initialize()

ComboBox1.Clear
ComboBox1.AddItem "Form Progress Listing Format"
ComboBox1.AddItem "Query Detail Listing Format Site"
ComboBox1.AddItem "Query Detail Listing Format Sponsor"
ComboBox1.AddItem "Subject Progress Listing Format"
ComboBox1.AddItem "Format Core Listing AE"
ComboBox1.AddItem "Format Core Listing Non AE"
ComboBox1.AddItem "AE Assessment Join Data"
ComboBox1.AddItem "15420 Ming's Comprehensive Form Summary Format"
ComboBox1.AddItem "03821 Ming's Comprehensive Form Summary Format"
ComboBox1.AddItem "Sheet Format (Left Align)"
ComboBox1.AddItem "Table Format (Left Align)"
ComboBox1.AddItem "Table Format (Wrap Text, Center Align)"


VeevaMacros.Label1.Caption = "Please select the appropriate automation from the list above." & vbNewLine & "For Table Format automation, please select the top left cell of the table before running automation." & vbNewLine & "For automation updates, please submit an Automation Request Form"
End Sub

Sub CommandButton1_Click()
    If ComboBox1.Value = "Form Progress Listing Format" Then
        Call VeevaReports.FormProgressListingFormat
    ElseIf ComboBox1.Value = "Query Detail Listing Format Site" Then
        Call VeevaReports.QueryDetailListingFormatSite
    ElseIf ComboBox1.Value = "Query Detail Listing Format Sponsor" Then
        Call VeevaReports.QueryDetailListingFormatSponsor
    ElseIf ComboBox1.Value = "Subject Progress Listing Format" Then
        Call VeevaReports.SubjectProgressListingFormat
    ElseIf ComboBox1.Value = "Format Core Listing AE" Then
        Call VeevaAE.FormatVeevaAE
    ElseIf ComboBox1.Value = "Format Core Listing Non AE" Then
        Call SafetyMacro.FormatNonAECRF
    ElseIf ComboBox1.Value = "AE Assessment Join Data" Then
        Call AssessmentJoinData.ProcessAssessmentJoinData
    ElseIf ComboBox1.Value = "15420 Ming's Comprehensive Form Summary Format" Then
        Call S15420.MingComprehensiveFormSummary
    ElseIf ComboBox1.Value = "03821 Ming's Comprehensive Form Summary Format" Then
        Call S03821.MingComprehensiveFormSummary
    ElseIf ComboBox1.Value = "Sheet Format (Left Align)" Then
        Call Main.OutFormat
    ElseIf ComboBox1.Value = "Table Format (Left Align)" Then
        Call HelperSub.FormatTable
    ElseIf ComboBox1.Value = "Table Format (Wrap Text, Center Align)" Then
        Call HelperSub.FormatTable2
    End If
    Unload VeevaMacros
End Sub

