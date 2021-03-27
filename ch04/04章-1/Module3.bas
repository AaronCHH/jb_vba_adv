Attribute VB_Name = "Module3"
Option Explicit
Option Base 1

'------------------------------------
'範例29
'保留陣列的值，變更元素數量
'------------------------------------

Sub PreserveSample()
    Dim myName() As String
    
    Dim r As Long, r2 As Long
    Dim i As Long

    '將「3年A班」工作表內的資料置入陣列變數中
    Worksheets("3年A班").Activate
    
    r = Range("A65536").End(xlUp).Row

    ReDim myName(r)
    
    For i = 1 To r
        myName(i) = Cells(i, 1).Value
    Next
        
    '將「3年B班」工作表內的資料新增到陣列變數中
    Worksheets("3年B班").Activate
    
    r2 = Range("A65536").End(xlUp).Row

    ReDim Preserve myName(r + r2)
    
    For i = r + 1 To r + r2
        myName(i) = Cells(i - r, 1).Value
    Next
    
    '將陣列變數的值輸出到即時運算視窗中
    For i = LBound(myName) To UBound(myName)
        Debug.Print myName(i)
    Next
End Sub

