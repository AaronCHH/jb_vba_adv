Attribute VB_Name = "Module1"
Option Explicit

'------------------
'範例26
'陣列變數的基本語法
'------------------

Sub OneWeek()
    Dim myWeek(6) As String
    Dim i As Integer
    
    myWeek(0) = "星期日"
    myWeek(1) = "星期一"
    myWeek(2) = "星期二"
    myWeek(3) = "星期三"
    myWeek(4) = "星期四"
    myWeek(5) = "星期五"
    myWeek(6) = "星期六"
    
    For i = 0 To 6
        Debug.Print myWeek(i)
    Next
End Sub
