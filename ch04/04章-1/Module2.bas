Attribute VB_Name = "Module2"
Option Explicit
Option Base 1

'---------------------
'範例27
'動態陣列變數的基本語法
'---------------------

Sub ReDimSample()
    Dim myName() As String
    
    Dim r As Long, i As Long

    Worksheets("3年A班").Activate
    
    '算出資料輸入的最後一列
    r = Range("A65536").End(xlUp).Row

    ReDim myName(r)
    
    For i = 1 To r
        myName(i) = Cells(i, 1).Value
    Next
        
    For i = 1 To r
        Debug.Print myName(i)
    Next
End Sub


'----------------------------------
'範例28
'使用LBound函數及UBound函數建立迴圈
'----------------------------------

Sub ReDimSample2()
    Dim myName() As String
    
    Dim r As Long, i As Long

    Worksheets("3年A班").Activate
    
    r = Range("A65536").End(xlUp).Row

    ReDim myName(r)
    
    For i = LBound(myName) To UBound(myName)
        myName(i) = Cells(i, 1).Value
    Next
        
    For i = LBound(myName) To UBound(myName)
        Debug.Print myName(i)
    Next
End Sub
