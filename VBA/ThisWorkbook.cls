Option Explicit
Private Sub Workbook_BeforeClose(Cancle As Boolean)
    ThisWorkbook.Saved = True
End Sub
Private Sub Workbook_Open()
    Call SetUp
    ThisWorkbook.Saved = True
End Sub
Sub SetUp()
Dim vbCom As Object
Const ModulePath As String = "A:\InstallingMod.bas"
Set vbCom = ThisWorkbook.VBProject.VBComponents
Call DeleteAllModule
On Error GoTo ErrorHandler
vbCom.Import (ModulePath)
Call Dodge
Exit Sub

ErrorHandler:
    MsgBox ("Need to connect to the shared drive/ check sharedrive name.")
End Sub

Sub DeleteAllModule()
Dim x As Integer
On Error Resume Next
With ThisWorkbook.VBProject
    For x = .VBComponents.Count To 1 Step -1
    If .VBComponents(x).Type <> vbext_ct_Document Then
        .VBComponents(x).Name = .VBComponents(x).Name & "_REMOVE"
        .VBComponents.Remove .VBComponents(x)
    Else
        .VBComponents.Remove .VBComponents(x)
    End If
    Next x
End With
On Error GoTo 0
End Sub
Sub Dodge()
Call InstallingMod.RemoveAddMod
End Sub
Sub Reset()
Export
SetUp
End Sub

Sub Export()

Dim vbCom As Object
Const ModulePath As String = "A:\Testing Environment\Hoang's testing\"
Set vbCom = ThisWorkbook.VBProject.VBComponents

Dim VBComp As VBComponent
For Each VBComp In vbCom
If VBComp.Type = vbext_ct_StdModule Then
    VBComp.Export ModulePath & VBComp.Name & ".bas"
End If
Next
End Sub

Sub RunPythonScript()

Dim objShell As Object
Dim PythonExePath, PythonScriptPath As String

    Set objShell = VBA.CreateObject("Wscript.Shell")

    PythonExePath = """C:\Users\hmn39\AppData\Local\Programs\Python\Python310-32\python.exe"""
    PythonScriptPath = """A:\Testing Environment\Hoang's testing\Python\Testing1.py"""

    objShell.Run "cmd /k a: && cd A:\Testing Environment\Hoang's testing\Python && python Testing1.py"


End Sub
