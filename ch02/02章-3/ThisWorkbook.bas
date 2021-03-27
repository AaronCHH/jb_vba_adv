Option Explicit

'------------------------------------------------------------
'範例16
'開啟特定工作表時執行事件程序
'------------------------------------------------------------

Private Sub Worksheet_Activate()
    Dim myWSName As String
    
    myWSName = ActiveSheet.Name
    MsgBox "請勿變更工作表" & myWSName & "的內容!"
End Sub
