Attribute VB_Name = "Module5"
Option Explicit

'--------------------------------
'範例33
'使用者自訂型態變數的基本使用方式
'--------------------------------

Type PersonalData
    PName As String
    PAge As Integer
    PDate As Date
End Type

Sub TypeSample()
    Dim myData As PersonalData
    
    myData.PName = "曹束昇"
    myData.PAge = 27
    myData.PDate = #4/1/1995#

    MsgBox "姓名：" & myData.PName & vbCrLf & _
        "年齡：" & myData.PAge & vbCrLf & _
        "到職日：" & myData.PDate
End Sub


'-----------------------------------
'範例34
'Static陳述式的基本使用方式
'-----------------------------------

Sub StaticSample()
    Static myNum As Integer
    
    myNum = myNum + 10
    MsgBox myNum
End Sub


'-----------------------------------
'範例35
'使用者自訂常數的基本使用方式
'-----------------------------------

Sub ConstSample()
    Const myBlue As Integer = 5
    
    Range("I11").Select
    Selection.Interior.ColorIndex = myBlue
End Sub

