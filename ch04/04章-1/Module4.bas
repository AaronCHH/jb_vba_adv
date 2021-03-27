Attribute VB_Name = "Module4"
Option Explicit
Option Base 1

'-------------------------
'範例30
'二維陣列的基本使用方式
'-------------------------

Sub SampleMatrix()
    Dim myData(3, 2) As String
    
    Dim i As Integer, j As Integer

    myData(1, 1) = "陳銘德"
    myData(1, 2) = "飛快汽車"

    myData(2, 1) = "梁銘鼎"
    myData(2, 2) = "鼎運資訊"

    myData(3, 1) = "彭凱堯"
    myData(3, 2) = "大富工程"

    For i = 1 To 3
        For j = 1 To 2
            Debug.Print myData(i, j)
        Next j
    Next i
End Sub


'--------------------------------------------------
'範例31
'在二維陣列中使用For Each...Next陳述式
'--------------------------------------------------

Sub SampleMatrix2()
    Dim myData(3, 2) As String
    
    Dim myVal As Variant

    myData(1, 1) = "陳銘德"
    myData(1, 2) = "飛快汽車"

    myData(2, 1) = "梁銘鼎"
    myData(2, 2) = "鼎運資訊"

    myData(3, 1) = "彭凱堯"
    myData(3, 2) = "大富工程"

    For Each myVal In myData
        Debug.Print myVal
    Next
End Sub


'--------------------------------------------------
'範例32
'將Variant型態變數設為儲存格範圍的值
'--------------------------------------------------

Sub RangeToVariant()
    Dim myData As Variant
    
    Dim r As Integer, c As Integer
    
    Worksheets("複本來源").Activate
    
    myData = Range("A1").CurrentRegion.Value

    
    r = UBound(myData, 1)
    c = UBound(myData, 2)
    
    Worksheets("複本").Activate
    
    Range(Cells(1, 1), Cells(r, c)).Value = myData
End Sub
