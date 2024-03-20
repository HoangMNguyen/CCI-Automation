Attribute VB_Name = "InstallingMod"
Sub RemoveAddMod(Optional ModulePath As String = "A:\VBA\")

Dim vbCom As Object
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
    Application.OnKey "+^{S}", "UserFormLaunch2"   'Formatting for individual CRF
    Application.OnKey "+^{V}", "VeevaForm"
    Application.OnKey "+^{P}", "VelosForm"

End Sub

Sub DeleteShortcut()
    Application.OnKey "+^{S}"
    Application.OnKey "+^{V}"
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

