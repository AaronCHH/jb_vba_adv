Attribute VB_Name = "Module1"
Option Explicit

'--------------------------------
'�d��6
'�H�ǧ}���覡�N�޼ƶǰe��Ƶ{��
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
'�d��7
'�H�ǭȪ��覡�N�޼ƶǰe��Ƶ{��
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
