Attribute VB_Name = "Module1"
Option Explicit

Sub StepMode()
    Dim i As Integer, j As Integer
    Dim mySum As Integer

    Worksheets("Sheet3").Activate
    Range("A1:B1").ClearContents
    
    i = 1
    j = 10
    mySum = F_AddNumber(i, j)
    
    Range("A1").Value = mySum
    Range("B1").Value = "¤¸"
End Sub

Function F_AddNumber(myMin, myMax)
    Dim k As Integer
    
    For k = myMin To myMax
        F_AddNumber = F_AddNumber + k
    Next k
End Function


Sub WriteA2()
    Dim myLen As Integer, i As Integer
    Dim myVal As Variant
    
    Worksheets("Sheet1").Activate
    
    myLen = Len(Range("A1").Value)
    
    For i = 1 To myLen
        myVal = Mid(Range("A1").Value, i, 1)
        If F_NumCheck(myVal) = True Then
            Range("A2").Value = Range("A2").Value & myVal
        End If
    Next
End Sub

Function F_NumCheck(v) As Boolean
    If v >= 0 And v <= 9 Then
        F_NumCheck = True
    Else
        F_NumCheck = False
    End If
End Function
