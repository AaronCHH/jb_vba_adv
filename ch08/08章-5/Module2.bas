Attribute VB_Name = "Module2"
Option Explicit

'-------------------------
'�d��91
'���J�W�q��
'-------------------------

Sub AddinInst()
    If AddIns("08��-5").Installed = False Then
        MsgBox "���J�W�q���u08��-5�v"
        AddIns("08��-5").Installed = True
    Else
        MsgBox "�W�q���u08��-5�v�w���J"
    End If
End Sub
