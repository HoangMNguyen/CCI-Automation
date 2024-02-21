Attribute VB_Name = "InstallingMod"
Sub RemoveAddMod()

Dim vbCom As Object
Const ModulePath As String = "A:\VBA Production\"
Set vbCom = ThisWorkbook.VBProject.VBComponents
On Error Resume Next
vbCom.Import (ModulePath & "Main.bas")
vbCom.Import (ModulePath & "S15CT055.bas")
vbCom.Import (ModulePath & "S18CT014.bas")
vbCom.Import (ModulePath & "S01817.bas")
vbCom.Import (ModulePath & "S01422.bas")
vbCom.Import (ModulePath & "S02916.bas")
vbCom.Import (ModulePath & "S03818.bas")
vbCom.Import (ModulePath & "S03821.bas")
vbCom.Import (ModulePath & "S12418.bas")
vbCom.Import (ModulePath & "S14217.bas")
vbCom.Import (ModulePath & "S32816.bas")
vbCom.Import (ModulePath & "S35418.bas")
vbCom.Import (ModulePath & "S15420.bas")
vbCom.Import (ModulePath & "S46417.bas")
vbCom.Import (ModulePath & "S19CT023.bas")
vbCom.Import (ModulePath & "ConmedAE.bas")
vbCom.Import (ModulePath & "VeevaAE.bas")
vbCom.Import (ModulePath & "FormattingUI.frm")
vbCom.Import (ModulePath & "VeevaMacros.frm")
vbCom.Import (ModulePath & "VelosMacros.frm")
vbCom.Import (ModulePath & "HelperSub.bas")
vbCom.Import (ModulePath & "SafetyMacro.bas")
vbCom.Import (ModulePath & "VeevaReports.bas")

Call CreatShortcut

End Sub


Sub CreatShortcut()
    Application.OnKey "+^{F}", "Main.QuickReportsFormStatusFormatSponsor" 'reformat quick report form status report
    Application.OnKey "+^{Q}", "Main.QuickReportQueryStatusFormat" 'reformat quick report patient form query status report
    Application.OnKey "+^{N}", "Main.QuickReportsFormStatusFormatSite" 'Form Status Report for Site (only WIP, INC)
    Application.OnKey "+^{S}", "UserFormLaunch2"   'Formatting for individual CRF
    Application.OnKey "+^{W}", "SafetyMacro.FormatNonAECRF"
    Application.OnKey "+^{V}", "VeevaForm"
    Application.OnKey "+^{P}", "VelosForm"
    Application.OnKey "+^{M}", "Main.VelosScheduleReview"
    Application.OnKey "+^{C}", "ConmedAE.FormatConmed" 'reformat single row conmed report
    Application.OnKey "+^{A}", "ConmedAE.FormatAE" 'reformat AE page and CCI AE report
    Application.OnKey "+^{Y}", "ConmedAE.FormatCRFs" 'reformat any CRF report


End Sub

Sub DeleteShortcut()
    Application.OnKey "+^{S}"
    Application.OnKey "+^{N}"
    Application.OnKey "+^{W}"
    Application.OnKey "+^{D}"
    Application.OnKey "+^{R}"
    Application.OnKey "+^{M}"
    Application.OnKey "+^{F}"
    Application.OnKey "+^{C}"
    Application.OnKey "+^{A}"
    Application.OnKey "+^{Y}"
    Application.OnKey "+^{V}"
    Application.OnKey "+^{Q}"
    Application.OnKey "+^{P}"
End Sub
Sub UserFormLaunch2()
    Load FormattingUI
    FormattingUI.StartUpPosition = 2
    FormattingUI.Show
    Call FormattingUI.UserForm_Initialize
End Sub
Sub VeevaForm()
    Load VeevaMacros
    VeevaMacros.StartUpPosition = 2
    VeevaMacros.Show
    Call VeevaMacros.UserForm_Initialize
End Sub

Sub VelosForm()
    Load VelosMacros
    VelosMacros.StartUpPosition = 2
    VelosMacros.Show
    Call VelosMacros.UserForm_Initialize
End Sub

