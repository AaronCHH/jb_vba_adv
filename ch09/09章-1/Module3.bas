Attribute VB_Name = "Module3"
Option Explicit

'-------------------------------------------------------
'�d��97
'�����HShell��ƱҰʪ����ε{�����e�i�J�ݾ����A
'-------------------------------------------------------

'�Ǧ^�J�s���󱱨�N�X�����
'�Ǧ^�ȡ@���\ = ���w�B�z��Open����N�X
' �@�@�@�@���� = NULL
Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Public Const PROCESS_QUERY_INFORMATION = &H400&


'�Ǧ^���w�B�z�������A�����
'�Ǧ^�ȡ@���\ = 0�H�~���ƭ�
'�@�@�@�@���� = 0
Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, _
    lpExitCode As Long) As Long
        
'�P�_���w���B�z�O�_����
'(�Y�|�������A�N�m�JSTILL_ACTIVE)
Public Const STATUS_PENDING = &H103&
Public Const STILL_ACTIVE = STATUS_PENDING


'�����}�Ҥ����󱱨�N�X�����
''�Ǧ^�ȡ@���\ = 0�H�~���ƭ�
'�@�@�@�@���� = 0
Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long


'�{��
Sub GetExitCodeProcess_Sample()
    Dim lngProcessId As Long    'Shell��ƪ��Ǧ^��
    Dim lngProcess As Long      'OpenProcess��ƪ��Ǧ^��
    Dim lngExitCode As Long     '�����{���X
    Dim rc As Long
     
    MsgBox "��[�T�w]���s��N�}�ҰO�ƥ��A�����O�ƥ����e�{�ǱN�B��Ȱ������A"
              
    '�ҰʰO�ƥ�
    lngProcessId = Shell("Notepad.exe", vbNormalFocus)
    
    '���o�HShell��ƱҰʪ����ε{�����B�z���󪺱���N�X
    lngProcess = OpenProcess(PROCESS_QUERY_INFORMATION, _
                            1, _
                            lngProcessId)
                            
    '�HGetExitCodeProcess���o�B�z���������A
    '�Y�Ұʪ����ε{���B��|�����������A�A�N��DoEvent�~���@�~�t�θ߰ݨ䪬�A
    Do
        rc = GetExitCodeProcess(lngProcess, lngExitCode)
        DoEvents
    Loop While lngExitCode = STILL_ACTIVE

    '�����}�Ҥ������󱱨�N�X
    rc = CloseHandle(lngProcess)

    MsgBox "�O�ƥ�����"
End Sub



