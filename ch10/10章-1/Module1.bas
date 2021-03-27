Attribute VB_Name = "Module1"
Option Explicit



'-----------------------------------------------------------
'範例99
'以AutoFilter方法篩選出「蛋捲」
'-----------------------------------------------------------

Sub AutoFilterSample1()
    Range("傳票").AutoFilter Field:=3, Criteria1:="蛋捲"
End Sub


'-----------------------------------------------------------
'範例100
'解除自動篩選，顯示所有資料
'-----------------------------------------------------------

Sub AutoFilterSample2()
    Range("傳票").AutoFilter
End Sub


'-----------------------------------------------------------
'範例101
'AutoFilterMode屬性的特性
'-----------------------------------------------------------

Sub AutoFilterModeSample()
    With ActiveSheet
        MsgBox "自動篩選模式為：" & .AutoFilterMode
        .AutoFilterMode = Not .AutoFilterMode
    End With
End Sub


'-----------------------------------------------------------
'範例102
'以AutoFilter方法篩選數值的尾碼
'-----------------------------------------------------------

Sub AutoFilterSample3()
    Dim myCell As Range
    Dim myCode As Variant
    
    myCode = Application.InputBox("請鍵入數字的尾碼")
    
    If myCode = False Then Exit Sub
    
    For Each myCell In Range("傳票").Offset(1).Resize(Range("傳票").Rows.Count - 1, 1)
        myCell.Value = "'" & myCell.Value
    Next
    
    Selection.AutoFilter Field:=1, Criteria1:="=*" & myCode
End Sub


'-----------------------------------------------------------
'範例103
'以AdvancedFilter方法變更清單內容
'-----------------------------------------------------------

Sub AdvancedFilterSample1()

    Range("傳票").AdvancedFilter xlFilterInPlace, Range("條件範圍")
    
End Sub


'-----------------------------------------------------------
'範例104
'將AdvancedFilter的篩選結果複製到其他地方
'-----------------------------------------------------------

Sub AdvancedFilterSample2()

    Range("傳票").AdvancedFilter xlFilterCopy, Range("條件範圍"), _
        Worksheets("篩選結果").Range("A2")

End Sub


'-----------------------------------------------------------
'還原清單
'-----------------------------------------------------------

Sub ResetList()
    
    If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData

End Sub



