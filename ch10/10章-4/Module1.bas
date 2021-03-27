Attribute VB_Name = "Module1"
Option Explicit


'-----------------------------------------------------------
'範例110
'以ListRow物件對列進行操作
'-----------------------------------------------------------

Sub ListRowSample()
    Dim myList As ListObject
    Dim i As Integer
    
    Set myList = ActiveSheet.ListObjects(1)
    
    For i = 1 To myList.ListRows.Count Step 2
        
        myList.ListRows(i).Range.Interior.ColorIndex = 32
    
    Next
End Sub


'-----------------------------------------------------------
'範例111
'依序顯示行標籤名稱
'-----------------------------------------------------------

Sub ListColumnSample1()
    Dim myList As ListObject
    Dim myCol As ListColumn
    
    Set myList = ActiveSheet.ListObjects(1)
    
    For Each myCol In myList.ListColumns
        
        MsgBox myCol.Name
    
    Next
End Sub


'-----------------------------------------------------------
'範例112
'選取「產品名稱」行的所有資料
'-----------------------------------------------------------

Sub ListColumnSample2()
    Dim myList As ListObject
    
    Set myList = ActiveSheet.ListObjects(1)
    
    myList.ListColumns("產品名稱").Range.Select
End Sub


'-----------------------------------------------------------
'範例113
'合計列
'-----------------------------------------------------------

Sub Syukeigyo()
    Dim myList As ListObject
    
    Set myList = ActiveSheet.ListObjects(1)
    
    myList.ShowTotals = True
    
    myList.TotalsRowRange.End(xlToRight).Select
End Sub

