Option Explicit

Sub myMacro1()
    MsgBox "執行巨集1", vbExclamation
End Sub
Sub myMacro2()
    MsgBox "執行巨集2", vbExclamation
End Sub

'-----------------------------
'範例92
'卸載增益集
'-----------------------------

Sub AddinUnInst()
    MsgBox "卸載增益集「08章-5」"
    AddIns("08章-5").Installed = False
End Sub
