Attribute VB_Name = "Module2"
Option Explicit

'-----------------------------
'�d��9
'��@�޼ƪ��ϥΪ̦ۭq���
'-----------------------------

Function MyMsg(Result As Variant) As String
    Select Case Result
        Case Is > 80
            MyMsg = "�u"
        Case Is > 60
            MyMsg = "�}"
        Case Is > 40
            MyMsg = "���ή�"
        Case Else
            MyMsg = "�ݧV�O"
    End Select
End Function

'-----------------------------
'�d��10
'�ƼƤ޼ƪ��ϥΪ̦ۭq���
'-----------------------------

Function MyFindS(r As Range, s As String) As Long
    Dim i As Long           '�ŦX�j�M���G���x�s���
    Dim myCell As Range

    For Each myCell In r
        If InStr(myCell.Value, s) > 0 Then
            i = i + 1
        End If
    Next
    
    MyFindS = i
End Function

'----------------------------------------------
'�d��11
'�[�`���w�x�s��d��ƭȪ��ϥΪ̦ۭq���
'----------------------------------------------

Function MySum(r As Range) As Double
    Dim myCell As Range

    For Each myCell In r
        MySum = Val(myCell.Value) + MySum
    Next
End Function

'----------------------------------------------
'�d��12
'�۰ʭ��йB�⪺�ϥΪ̦ۭq���
'----------------------------------------------

Function MyKaihi(n As Integer) As Long
    Application.Volatile
    MyKaihi = Range("A2").Value * n
End Function

