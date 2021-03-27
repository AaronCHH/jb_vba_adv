Attribute VB_Name = "Module1"
Option Explicit

'--------------------------------
'範例55
'隱藏所有命令列
'--------------------------------

Sub UnVisibleAllCmdBars()
    Dim myCB As CommandBar

    On Error Resume Next        '若發生錯誤，程序繼續執行
    
    For Each myCB In Application.CommandBars
        myCB.Visible = False
    Next myCB
    
    On Error GoTo 0             '停用錯誤處理副程序
    
    Application.CommandBars("Worksheet Menu Bar").Enabled = False
    
    MsgBox "按[確定]將重新顯示被隱藏的所有命令列：" & Chr(13) & Chr(13) & _
        "‧工作表功能表列" & Chr(13) & _
        "‧[一般]工具列" & Chr(13) & _
        "‧[格式]工具列"
    
    With Application
        .CommandBars("Worksheet Menu Bar").Enabled = True
        .CommandBars("Standard").Visible = True
        .CommandBars("Formatting").Visible = True
    End With
End Sub

