Option Explicit

Sub myMacro1()
    MsgBox "���楨��1", vbExclamation
End Sub
Sub myMacro2()
    MsgBox "���楨��2", vbExclamation
End Sub

'-----------------------------
'�d��92
'�����W�q��
'-----------------------------

Sub AddinUnInst()
    MsgBox "�����W�q���u08��-5�v"
    AddIns("08��-5").Installed = False
End Sub
