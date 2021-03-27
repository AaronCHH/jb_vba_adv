Attribute VB_Name = "Module1"
'-----------------------------------------
'範例1
'從主程序呼叫副程序
'(排序→篩選→列印→停用自動篩選)
'-----------------------------------------

Sub PrintMember()
    SortData        '副程序(排序)
    
    SelectData      '副程序(篩選)
    
    PrintData       '副程序(列印)
    
    KaijoFilter     '副程序(停用自動篩選)
End Sub

'-----------
'副程序
'-----------

'1.排序
Sub SortData()
    Range("A1").Sort Key1:=Range("D3"), Order1:=xlAscending, _
        Key2:=Range("F3"), Order2:=xlAscending, Header:=xlGuess
End Sub

'2.篩選
Sub SelectData()
    Range("A1").AutoFilter Field:=9, Criteria1:=">=2003/1/1", _
        Operator:=xlAnd, Criteria2:="<=2006/12/31"
End Sub

'3.列印資料
Sub PrintData()
    Range("A1").CurrentRegion.Offset(1).Select
    Selection.Resize(Selection.Rows.Count - 1).Select
    ActiveSheet.PageSetup.PrintArea = Selection.Address
    ActiveSheet.PrintOut
End Sub

'4.停用自動篩選
Sub KaijoFilter()
    Range("A1").Select
    Selection.AutoFilter
End Sub
