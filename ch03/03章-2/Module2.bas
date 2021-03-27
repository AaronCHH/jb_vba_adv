Attribute VB_Name = "Module2"
Option Explicit

Sub LocalWindow()
    Dim myWS As Worksheet
    Dim myName(4) As String
    Dim i As Integer
    
    Set myWS = Worksheets(2)
    
    For i = 0 To 4
        myName(i) = myWS.Cells(i + 1, 1).Value
    Next
End Sub

