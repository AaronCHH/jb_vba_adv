Attribute VB_Name = "Module1"
Option Explicit

'--------------------------------
'�d��55
'���éҦ��R�O�C
'--------------------------------

Sub UnVisibleAllCmdBars()
    Dim myCB As CommandBar

    On Error Resume Next        '�Y�o�Ϳ��~�A�{���~�����
    
    For Each myCB In Application.CommandBars
        myCB.Visible = False
    Next myCB
    
    On Error GoTo 0             '���ο��~�B�z�Ƶ{��
    
    Application.CommandBars("Worksheet Menu Bar").Enabled = False
    
    MsgBox "��[�T�w]�N���s��ܳQ���ê��Ҧ��R�O�C�G" & Chr(13) & Chr(13) & _
        "�E�u�@��\���C" & Chr(13) & _
        "�E[�@��]�u��C" & Chr(13) & _
        "�E[�榡]�u��C"
    
    With Application
        .CommandBars("Worksheet Menu Bar").Enabled = True
        .CommandBars("Standard").Visible = True
        .CommandBars("Formatting").Visible = True
    End With
End Sub

