Attribute VB_Name = "Module2"
'-------------------------------
'�d��2
'�I�s�ϥΤ޼ƪ��Ƶ{��
'-------------------------------

Sub SortMember()
    Dim myRowNo1 As Integer, myRowNo2 As Integer

    Worksheets("�|���W�U").AutoFilterMode = False
    
    myRowNo1 = Application.InputBox("�п�J���ƧǪ����(�ĴX��)", _
        "�����Ĥ@����")
    
    If myRowNo1 < 1 Or myRowNo1 > 10 Then Exit Sub
    
    myRowNo2 = Application.InputBox("�п�J���ƧǪ����(�ĴX��)", _
        "�����ĤG����")
    
    If myRowNo2 < 1 Or myRowNo2 > 10 Then Exit Sub
        
    RunSort myRowNo1, myRowNo2
End Sub

Sub RunSort(myR1 As Integer, myR2 As Integer)
    Range("A1").Sort Key1:=Cells(3, myR1), Order1:=xlAscending, _
        Key2:=Cells(3, myR2), Order2:=xlAscending, Header:=xlGuess

    'MsgBox myR1
    'MsgBox myR2
End Sub
