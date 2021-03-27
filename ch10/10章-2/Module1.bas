Attribute VB_Name = "Module1"
Option Explicit


'-----------------------------------------------------------
'範例105
'以AdvancedFilter方法重複篩選資料
'-----------------------------------------------------------

Sub AdvancedFilterSample3()
    Dim myRange As Range
    
    Range("I5:I15").Clear
    
    Set myRange = Range("Data").Offset(, 1).Resize(Range("Data").Rows.Count, 1)
    
    myRange.AdvancedFilter xlFilterCopy, , Range("I5"), True
End Sub


'-----------------------------------------------------------
'範例106
'以AdvancedFilter方法建立個別顧客銷售資料
'-----------------------------------------------------------

Sub AdvancedFilterSample4()
    Dim myRange As Range
    Dim myJoken As Range
    Dim mySheet As Worksheet
    Dim i As Integer
    
    Range("I5:I15").Clear
    
    Set myRange = Range("Data").Offset(, 1).Resize(Range("Data").Rows.Count, 1)
    
    myRange.AdvancedFilter xlFilterCopy, , Range("I5"), True
    
    Set myJoken = Range("I5").CurrentRegion
    
    For i = 1 To myJoken.Rows.Count - 1
        
        Set mySheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        
        mySheet.Name = myJoken.Cells(2, 1).Value
        
        Range("Data").AdvancedFilter xlFilterCopy, _
            myJoken.Rows("1:2"), mySheet.Range("A2")
        
        myJoken.Rows(2).Delete xlShiftUp
    Next
    
    myJoken.Clear
End Sub


