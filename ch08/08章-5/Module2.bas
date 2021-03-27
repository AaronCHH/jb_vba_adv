Attribute VB_Name = "Module2"
Option Explicit

'-------------------------
'範例91
'載入增益集
'-------------------------

Sub AddinInst()
    If AddIns("08章-5").Installed = False Then
        MsgBox "載入增益集「08章-5」"
        AddIns("08章-5").Installed = True
    Else
        MsgBox "增益集「08章-5」已載入"
    End If
End Sub
