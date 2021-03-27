Option Explicit

'--------------------------------------------------------
'變更儲存格的捷徑功能表
'--------------------------------------------------------

Private Sub Workbook_Open()
    Dim myCBCtrl As CommandBarButton
        
    Set myCBCtrl = Application.CommandBars("Cell").Controls.Add _
        (Type:=msoControlButton, Id:=872, Before:=8, Temporary:=True)
    myCBCtrl.Caption = "清除格式(&F)"
End Sub

'--------------------------------------------------------
'將變更過的捷徑功能表還原回預設值
'--------------------------------------------------------

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.CommandBars("Cell").Reset
End Sub
