Attribute VB_Name = "Module1"
Option Explicit

'------------------------------------
'�d��21
'On Error GoTo���z���d��
'------------------------------------

Sub TrapSample()
    
    '�ҥ�Error Trap�\��
    On Error GoTo HandleErr
    
    '�w���|�o�Ϳ��~�����z��
    ActiveWorkbook.Charts(1).Activate
    
    '�S���o�Ϳ��~���B�z�{��
    ActiveChart.SizeWithWindow = True
    MsgBox "���槹���A�S���o�Ϳ��~�C"
    Exit Sub
    
HandleErr:
    MsgBox "�Ϫ�u�@���s�b�C"
End Sub

'-------------------------------------------
'�d��22
'On Error Resume Next���z���d��
'-------------------------------------------

Sub TrapSample2()
    Dim myRange As Range
    Dim myPrompt As String, myTitle As String
    
    Worksheets("Sheet3").Activate
    Cells.Clear
    
    myPrompt = "�Q������x�s��d��N��J�uABC�v" & vbCr & _
        "�Шϥηƹ�����x�s��d��C"
    myTitle = "��J�x�s��d��"
    
    '�ҥ�Error Trap�\��
    On Error Resume Next
    
    '�w���|�o�Ϳ��~�����z��
    Set myRange = Application.InputBox(Prompt:=myPrompt, _
        Title:=myTitle, Type:=8)
    
    '�P�_�e�������z���O�_�o�Ϳ��~
    If myRange Is Nothing Then Exit Sub

    myRange.Value = "ABC"
End Sub

'-----------------------------
'�d��23
'�d�߿��~�N�X�ο��~�T��
'-----------------------------

Sub DisplayErr()
    Dim myMsg As String
    
    Worksheets("Sheet4").Activate
    
    On Error GoTo HandleErr
    
    Range("B3").Value = Range("B1").Value / Range("B2").Value
    
    Exit Sub

HandleErr:
    myMsg = "���~�N�X�G" & Err.Number & vbCrLf & _
        "���~�T���G" & Err.Description
    MsgBox myMsg

    Range("B3").Value = 0
End Sub

'--------------------------------
'�d��24
'�̾ڿ��~�����P�_����
'--------------------------------

Sub OpenFile()
    Dim myFD As Variant, myFN As Variant
    Dim myPrompt As String, myMsg As String
    Dim myBuf As String
    
    MsgBox "�Цb�Ϻо����ڥؿ��U�ǳƤ@�Ӥ�r�ɮסA" _
        & vbCr & "�������ɮצW�١C"
    
InputFD:
    myPrompt = "�п�J�Ϻо��N���G"
    myFD = Application.InputBox(Prompt:=myPrompt, Default:="A")
    If VarType(myFD) <> vbString Then Exit Sub
    
InputFN:
    myPrompt = "�п�J�ɮצW�١G"
    myFN = Application.InputBox(Prompt:=myPrompt)
    If VarType(myFN) <> vbString Then Exit Sub
    
    On Error GoTo HandleErr
    
    Open myFD & ":\" & myFN For Input As #1
    
    Do Until EOF(1)
        Line Input #1, myBuf
    Loop
    
    MsgBox "�B�z�����A�S���o�Ϳ��~�C"
    Close #1
    
    Exit Sub

HandleErr:
    Select Case Err.Number
        Case 53                 '�䤣���ɮ�
            MsgBox Err.Description & vbCr & _
                 "�Э��s��J�ɮצW�١G"
            Resume InputFN
            
        Case 55                 '�ɮפw�}��
            MsgBox Err.Description
            Resume Next
        
        Case 68, 75, 76         '�g��]�ƵL�k�ϥ�
            MsgBox Err.Description & vbCr & vbCr & _
                "���w���ϺеL�ġA" & vbCr & _
                "�ЦA��J�ϺϥN���G"
            Resume InputFD
        
        Case 52, 71             '�ϺЩ|���N��
            myMsg = Err.Description & vbCr & _
                "�n���J�Ϥ��~��ܡH"
            If MsgBox(myMsg, vbExclamation + vbYesNo) = vbYes Then
                Resume
            Else
                Exit Sub
            End If
    End Select
End Sub

'------------------------------------
'�d��25
'�N�ȿ�X��Y�ɹB�����
'------------------------------------

Sub OutToWindow()
    Dim n As Integer, m As Integer
    
    For n = 1 To 10
        m = 2 ^ n
        Debug.Print "m=" & m        '�}�ҧY�ɹB�����
    Next
    
    'Stop
End Sub

