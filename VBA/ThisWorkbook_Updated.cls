VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
Const ModulePath As String = "A:\VBA Production\"
Const InstallingMod As String = "InstallingMod_Updated.bas"
Set vbCom = ThisWorkbook.VBProject.VBComponents
Call DeleteAllModule
On Error GoTo ErrorHandler
vbCom.Import (ModulePath & InstallingMod)
Call Dodge(ModulePath)
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
Sub Dodge(ModulePath As String)
Call InstallingMod.RemoveAddMod(ModulePath)
End Sub
Sub Reset()
Export
SetUp
End Sub