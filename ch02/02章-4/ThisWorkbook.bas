Option Explicit

'--------------------------------------------------------------
'範例17
'開啟不特定工作表時執行事件程序
'--------------------------------------------------------------

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    MsgBox "請勿變更工作表" & Sh.Name & "的內容！"
End Sub