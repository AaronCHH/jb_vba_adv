Attribute VB_Name = "Module1"
Option Explicit

'------------------
'�d��26
'�}�C�ܼƪ��򥻻y�k
'------------------

Sub OneWeek()
    Dim myWeek(6) As String
    Dim i As Integer
    
    myWeek(0) = "�P����"
    myWeek(1) = "�P���@"
    myWeek(2) = "�P���G"
    myWeek(3) = "�P���T"
    myWeek(4) = "�P���|"
    myWeek(5) = "�P����"
    myWeek(6) = "�P����"
    
    For i = 0 To 6
        Debug.Print myWeek(i)
    Next
End Sub
