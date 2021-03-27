Attribute VB_Name = "Module1"
Option Explicit

'--------------------------------
'範例6
'以傳址的方式將引數傳送到副程序
'--------------------------------

Sub SansyoWatashi()
    Dim myNumber As Integer
    
    myNumber = 1
    
    ChangeNumber1 myNumber
    
    MsgBox myNumber
End Sub

Sub ChangeNumber1(ByRef n As Integer)
     
    n = 2

End Sub


'--------------------------------
'範例7
'以傳值的方式將引數傳送到副程序
'--------------------------------

Sub AtaiWatashi()
    Dim myNumber As Integer
    
    myNumber = 1
    
    ChangeNumber2 myNumber
    
    MsgBox myNumber
End Sub

Sub ChangeNumber2(ByVal n As Integer)
     
    n = 2

End Sub
