Option Explicit

'建立工作列
Sub AddCmdBar2()
    Dim myCB As CommandBar, myCBCtrl As CommandBarControl
    
    On Error Resume Next
    CommandBars("MyMacro4").Delete
    On Error GoTo 0
    
    Set myCB = Application.CommandBars.Add(Name:="MyMacro4")
       
    '建立切換按鈕
    Set myCBCtrl = myCB.Controls.Add(Type:=msoControlButton, Before:=1)
    
    With myCBCtrl
        .Caption = "ToggleBtn"
        .FaceId = 113
        .OnAction = "CBCtrlState1"
    End With
           
'-----建立[CheckMark]功能表-----
    Set myCBCtrl = myCB.Controls.Add(Type:=msoControlPopup, Before:=2)
    myCBCtrl.Caption = "CheckMark"
        
    '建立[CheckMark]→[CheckMarkOnOff]命令
    Set myCBCtrl = myCB.Controls("CheckMark").Controls _
        .Add(Type:=msoControlButton, Before:=1)
    myCBCtrl.Caption = "CheckMarkOnOff"
    myCBCtrl.OnAction = "CBCtrlState2"
        
'-----建立[刪除]功能表-----
    Set myCBCtrl = myCB.Controls.Add(Type:=msoControlPopup, Before:=3)
    myCBCtrl.Caption = "刪除(&D)"
    myCBCtrl.BeginGroup = True          '???????
    
    '建立[刪除]→[工具列]命令
    Set myCBCtrl = myCB.Controls("刪除(&D)").Controls _
        .Add(Type:=msoControlButton, Before:=1)
    myCBCtrl.Caption = "工具列(&T)"
    myCBCtrl.OnAction = "DeleteCmdBar"

    myCB.Visible = True

    MsgBox "按下按鈕時會切換開/關的狀態，" & Chr(13) & _
        "若點功能表，也會切換開/關的狀態。" & Chr(13) & _
         Chr(13) & "若需刪除此工具列，" & Chr(13) & _
        "請點[刪除]→[工具列]。"
End Sub


'-------------------------------------------
'範例57
'切換按鈕的狀態
'-------------------------------------------

Private Sub CBCtrlState1()
    Dim myCBCtrl As CommandBarButton
    
    Set myCBCtrl = CommandBars("MyMacro4").Controls("ToggleBtn")
    
    If myCBCtrl.State = msoButtonDown Then
         myCBCtrl.State = msoButtonUp
         'myCBCtrl.State = msoButtonMixed
         MsgBox "按鈕的狀態為「關」"
    Else
         myCBCtrl.State = msoButtonDown
         MsgBox "按鈕的狀態為「開」"
      End If
End Sub

'---------------------------------------------
'範例57
'切換功能表的狀態
'---------------------------------------------

Private Sub CBCtrlState2()
    Dim myCBCtrl As CommandBarButton
    
    Set myCBCtrl = CommandBars("MyMacro4").Controls("CheckMark") _
        .Controls("CheckMarkOnOff")
    
    If myCBCtrl.State = msoButtonDown Then
         myCBCtrl.State = msoButtonUp
         'myCBCtrl.State = msoButtonMixed
         MsgBox "Check Mark的狀態為「關」"
    Else
         myCBCtrl.State = msoButtonDown
         MsgBox "Check Mark的狀態為「開」"
      End If
End Sub


'刪除工具列
Private Sub DeleteCmdBar()
    MsgBox "刪除工具列「MyMacro4」"
    Application.CommandBars("MyMacro4").Delete
End Sub


