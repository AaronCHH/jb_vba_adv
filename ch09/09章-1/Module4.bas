Attribute VB_Name = "Module4"
Option Explicit

'-------------------------------------------------------
'�d��98
'�קK���ƱҰ����ε{��
'-------------------------------------------------------


'�qClass�W�٩�Caption���o��������N�X�����
'�Ǧ^�ȡ@���\ = ��������N�X
' �@�@�@�@���� = NULL
Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" _
    (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

'�{��
Sub FindWindow_Sample()
    Dim strClassName As String  'Class�W��
    Dim rc As Long
    
    Dim lngProcessId As Long    'Shell��ƪ��Ǧ^��
    
    '���wClass�W��
    strClassName = "SciCalc"
        
    '���o�p��L����������N�X
    rc = FindWindow(strClassName, _
                    vbNullString)
                    
    '���o��������N�X���|���Ұ�
    If rc <> 0& Then
        MsgBox "�p��L�w�}��"
        Exit Sub
    End If
        
    '�Ұʤp��L
    lngProcessId = Shell("Calc.exe", vbNormalFocus)
End Sub




