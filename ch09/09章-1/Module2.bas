Attribute VB_Name = "Module2"
Option Explicit
  
'------------------------------
'�d��96
'����O�ƥ�
'------------------------------

Sub ControlNotePad()
    Dim myPath As String
    Dim myID As Double             'Shell��ƪ��Ǧ^��

    myPath = ActiveWorkbook.Path & "\"

    '�}�ҰO�ƥ�
    myID = Shell("Notepad.exe", vbNormalFocus)
    
'<�Y�L�k�}�ҰO�ƥ��A�Х[�W�O�ƥ��{�������|>
    'myID = Shell("C:\Windows\Notepad.exe", vbNormalFocus)

    '��[Alt]+[F]+[O]��ǰe�}��[�}������]��ܤ�����R�O
    SendKeys "%FO", True

    '���w�öǰe�}���ɮצW�٪�����N�X
    SendKeys myPath & "Report.txt", True

    '��[ENTER]��ǰe�}���ɮת��R�O
    SendKeys "{ENTER}", True

    '�ƻs�x�s�檺���e
    Worksheets("���y�c��q").Range("���y�c��q�w��").Copy
    
    '�}�ҰO�ƥ�
    AppActivate myID
    
    '��[Ctrl]+{V]��ǰe�K�W�ƻs���e���R�O
    SendKeys "^V", True
    
    '�Ѱ��ƻs�Ҧ�
    Application.CutCopyMode = False
End Sub




