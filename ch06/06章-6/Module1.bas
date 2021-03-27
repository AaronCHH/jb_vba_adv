Option Explicit

'�إߤu�@�C
Sub AddCmdBar2()
    Dim myCB As CommandBar, myCBCtrl As CommandBarControl
    
    On Error Resume Next
    CommandBars("MyMacro4").Delete
    On Error GoTo 0
    
    Set myCB = Application.CommandBars.Add(Name:="MyMacro4")
       
    '�إߤ������s
    Set myCBCtrl = myCB.Controls.Add(Type:=msoControlButton, Before:=1)
    
    With myCBCtrl
        .Caption = "ToggleBtn"
        .FaceId = 113
        .OnAction = "CBCtrlState1"
    End With
           
'-----�إ�[CheckMark]�\���-----
    Set myCBCtrl = myCB.Controls.Add(Type:=msoControlPopup, Before:=2)
    myCBCtrl.Caption = "CheckMark"
        
    '�إ�[CheckMark]��[CheckMarkOnOff]�R�O
    Set myCBCtrl = myCB.Controls("CheckMark").Controls _
        .Add(Type:=msoControlButton, Before:=1)
    myCBCtrl.Caption = "CheckMarkOnOff"
    myCBCtrl.OnAction = "CBCtrlState2"
        
'-----�إ�[�R��]�\���-----
    Set myCBCtrl = myCB.Controls.Add(Type:=msoControlPopup, Before:=3)
    myCBCtrl.Caption = "�R��(&D)"
    myCBCtrl.BeginGroup = True          '???????
    
    '�إ�[�R��]��[�u��C]�R�O
    Set myCBCtrl = myCB.Controls("�R��(&D)").Controls _
        .Add(Type:=msoControlButton, Before:=1)
    myCBCtrl.Caption = "�u��C(&T)"
    myCBCtrl.OnAction = "DeleteCmdBar"

    myCB.Visible = True

    MsgBox "���U���s�ɷ|�����}/�������A�A" & Chr(13) & _
        "�Y�I�\���A�]�|�����}/�������A�C" & Chr(13) & _
         Chr(13) & "�Y�ݧR�����u��C�A" & Chr(13) & _
        "���I[�R��]��[�u��C]�C"
End Sub


'-------------------------------------------
'�d��57
'�������s�����A
'-------------------------------------------

Private Sub CBCtrlState1()
    Dim myCBCtrl As CommandBarButton
    
    Set myCBCtrl = CommandBars("MyMacro4").Controls("ToggleBtn")
    
    If myCBCtrl.State = msoButtonDown Then
         myCBCtrl.State = msoButtonUp
         'myCBCtrl.State = msoButtonMixed
         MsgBox "���s�����A���u���v"
    Else
         myCBCtrl.State = msoButtonDown
         MsgBox "���s�����A���u�}�v"
      End If
End Sub

'---------------------------------------------
'�d��57
'�����\������A
'---------------------------------------------

Private Sub CBCtrlState2()
    Dim myCBCtrl As CommandBarButton
    
    Set myCBCtrl = CommandBars("MyMacro4").Controls("CheckMark") _
        .Controls("CheckMarkOnOff")
    
    If myCBCtrl.State = msoButtonDown Then
         myCBCtrl.State = msoButtonUp
         'myCBCtrl.State = msoButtonMixed
         MsgBox "Check Mark�����A���u���v"
    Else
         myCBCtrl.State = msoButtonDown
         MsgBox "Check Mark�����A���u�}�v"
      End If
End Sub


'�R���u��C
Private Sub DeleteCmdBar()
    MsgBox "�R���u��C�uMyMacro4�v"
    Application.CommandBars("MyMacro4").Delete
End Sub


