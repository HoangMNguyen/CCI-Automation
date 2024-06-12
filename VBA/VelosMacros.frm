VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VelosMacros 
   Caption         =   "PennCTMS Automation"
   ClientHeight    =   1890
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   6360
   OleObjectBlob   =   "VelosMacros.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VelosMacros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Sub UserForm_Initialize()

ComboBox1.Clear
ComboBox1.AddItem "Quick Report Query Status Format"
ComboBox1.AddItem "Quick Reports Form Status Format Site"
ComboBox1.AddItem "Quick Reports Form Status Format Sponsor"
ComboBox1.AddItem "Velos/PennCTMS Schedule Review"
ComboBox1.AddItem "Format Conmed"
ComboBox1.AddItem "Format AE"
ComboBox1.AddItem "Format PDAE"
ComboBox1.AddItem "Format Quick AE"
ComboBox1.AddItem "Format Quick ConMed"
ComboBox1.AddItem "Format CRFs"
ComboBox1.AddItem "Sheet Format (Left Align)"
ComboBox1.AddItem "Table Format (Left Align)"
ComboBox1.AddItem "Table Format (Wrap Text, Center Align)"


VelosMacros.Label1.Caption = "Please select the appropriate automation from the list above." & vbNewLine & "For Table Format automation, please select the top left cell of the table before running automation." & vbNewLine & "For automation updates, please submit an Automation Request Form."
End Sub

Sub CommandButton1_Click()
    If ComboBox1.Value = "Quick Report Query Status Format" Then
        Call Main.QuickReportQueryStatusFormat
    ElseIf ComboBox1.Value = "Quick Reports Form Status Format Site" Then
        Call Main.QuickReportsFormStatusFormatSite
    ElseIf ComboBox1.Value = "Quick Reports Form Status Format Sponsor" Then
        Call Main.QuickReportsFormStatusFormatSponsor
    ElseIf ComboBox1.Value = "Velos/PennCTMS Schedule Review" Then
        Call Main.VelosScheduleReview
    ElseIf ComboBox1.Value = "Format Conmed" Then
        Call ConmedAE.FormatConmed
    ElseIf ComboBox1.Value = "Format AE" Then
        Call ConmedAE.FormatAE
    ElseIf ComboBox1.Value = "Format PDAE" Then
        Call ConmedAE.FormatPDAE
    ElseIf ComboBox1.Value = "Format CRFs" Then
        Call ConmedAE.FormatCRFs
    ElseIf ComboBox1.Value = "Format Quick AE" Then
        Call ConmedAE.FormatQuickAE
    ElseIf ComboBox1.Value = "Format Quick ConMed" Then
        Call ConmedAE.FormatQuickConMed
    ElseIf ComboBox1.Value = "Sheet Format (Left Align)" Then
        Call Main.OutFormat
    ElseIf ComboBox1.Value = "Table Format (Left Align)" Then
        Call HelperSub.FormatTable
    ElseIf ComboBox1.Value = "Table Format (Wrap Text, Center Align)" Then
        Call HelperSub.FormatTable2
    End If
    Unload VelosMacros
End Sub


