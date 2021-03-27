Attribute VB_Name = "Module1"
Option Explicit

'-----------------------------
'範例8
'建立Function程序
'-----------------------------

Sub TestResult()
    Range("A1").Select
    
    MsgBox TestMsg
End Sub

Function TestMsg() As String
    Select Case ActiveCell.Value
        Case Is > 80
            TestMsg = "優"
        Case Is > 60
            TestMsg = "良"
        Case Is > 40
            TestMsg = "不及格"
        Case Else
            TestMsg = "需努力"
    End Select
End Function
