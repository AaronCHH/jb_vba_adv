Attribute VB_Name = "Module2"
'-------------------------------
'範例2
'呼叫使用引數的副程序
'-------------------------------

Sub SortMember()
    Dim myRowNo1 As Integer, myRowNo2 As Integer

    Worksheets("會員名冊").AutoFilterMode = False
    
    myRowNo1 = Application.InputBox("請輸入欲排序的行數(第幾行)", _
        "此為第一順位")
    
    If myRowNo1 < 1 Or myRowNo1 > 10 Then Exit Sub
    
    myRowNo2 = Application.InputBox("請輸入欲排序的行數(第幾行)", _
        "此為第二順位")
    
    If myRowNo2 < 1 Or myRowNo2 > 10 Then Exit Sub
        
    RunSort myRowNo1, myRowNo2
End Sub

Sub RunSort(myR1 As Integer, myR2 As Integer)
    Range("A1").Sort Key1:=Cells(3, myR1), Order1:=xlAscending, _
        Key2:=Cells(3, myR2), Order2:=xlAscending, Header:=xlGuess

    'MsgBox myR1
    'MsgBox myR2
End Sub
