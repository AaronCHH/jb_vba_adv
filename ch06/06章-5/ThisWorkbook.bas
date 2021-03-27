Option Explicit

'-----------------------
'範例56
'停用功能表
'-----------------------

Private Sub Workbook_Open()
    Dim myCB As CommandBar, myCBCtrl As CommandBarControl
    
    Set myCB = Application.CommandBars("Worksheet Menu Bar")
    
    myCB.Controls("格式(&O)").Enabled = False
    myCB.Controls("工具(&T)").Controls("選項(&O)...").Enabled = False
End Sub
  
'回復功能表的狀態
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.CommandBars("Worksheet Menu Bar").Reset
End Sub

