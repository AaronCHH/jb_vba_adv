Attribute VB_Name = "Module1"
Option Explicit

'-----------------------------
'�d��8
'�إ�Function�{��
'-----------------------------

Sub TestResult()
    Range("A1").Select
    
    MsgBox TestMsg
End Sub

Function TestMsg() As String
    Select Case ActiveCell.Value
        Case Is > 80
            TestMsg = "�u"
        Case Is > 60
            TestMsg = "�}"
        Case Is > 40
            TestMsg = "���ή�"
        Case Else
            TestMsg = "�ݧV�O"
    End Select
End Function
