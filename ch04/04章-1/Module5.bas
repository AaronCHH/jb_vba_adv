Attribute VB_Name = "Module5"
Option Explicit

'--------------------------------
'�d��33
'�ϥΪ̦ۭq���A�ܼƪ��򥻨ϥΤ覡
'--------------------------------

Type PersonalData
    PName As String
    PAge As Integer
    PDate As Date
End Type

Sub TypeSample()
    Dim myData As PersonalData
    
    myData.PName = "����@"
    myData.PAge = 27
    myData.PDate = #4/1/1995#

    MsgBox "�m�W�G" & myData.PName & vbCrLf & _
        "�~�֡G" & myData.PAge & vbCrLf & _
        "��¾��G" & myData.PDate
End Sub


'-----------------------------------
'�d��34
'Static���z�����򥻨ϥΤ覡
'-----------------------------------

Sub StaticSample()
    Static myNum As Integer
    
    myNum = myNum + 10
    MsgBox myNum
End Sub


'-----------------------------------
'�d��35
'�ϥΪ̦ۭq�`�ƪ��򥻨ϥΤ覡
'-----------------------------------

Sub ConstSample()
    Const myBlue As Integer = 5
    
    Range("I11").Select
    Selection.Interior.ColorIndex = myBlue
End Sub

