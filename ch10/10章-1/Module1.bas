Attribute VB_Name = "Module1"
Option Explicit



'-----------------------------------------------------------
'�d��99
'�HAutoFilter��k�z��X�u�J���v
'-----------------------------------------------------------

Sub AutoFilterSample1()
    Range("�ǲ�").AutoFilter Field:=3, Criteria1:="�J��"
End Sub


'-----------------------------------------------------------
'�d��100
'�Ѱ��۰ʿz��A��ܩҦ����
'-----------------------------------------------------------

Sub AutoFilterSample2()
    Range("�ǲ�").AutoFilter
End Sub


'-----------------------------------------------------------
'�d��101
'AutoFilterMode�ݩʪ��S��
'-----------------------------------------------------------

Sub AutoFilterModeSample()
    With ActiveSheet
        MsgBox "�۰ʿz��Ҧ����G" & .AutoFilterMode
        .AutoFilterMode = Not .AutoFilterMode
    End With
End Sub


'-----------------------------------------------------------
'�d��102
'�HAutoFilter��k�z��ƭȪ����X
'-----------------------------------------------------------

Sub AutoFilterSample3()
    Dim myCell As Range
    Dim myCode As Variant
    
    myCode = Application.InputBox("����J�Ʀr�����X")
    
    If myCode = False Then Exit Sub
    
    For Each myCell In Range("�ǲ�").Offset(1).Resize(Range("�ǲ�").Rows.Count - 1, 1)
        myCell.Value = "'" & myCell.Value
    Next
    
    Selection.AutoFilter Field:=1, Criteria1:="=*" & myCode
End Sub


'-----------------------------------------------------------
'�d��103
'�HAdvancedFilter��k�ܧ�M�椺�e
'-----------------------------------------------------------

Sub AdvancedFilterSample1()

    Range("�ǲ�").AdvancedFilter xlFilterInPlace, Range("����d��")
    
End Sub


'-----------------------------------------------------------
'�d��104
'�NAdvancedFilter���z�ﵲ�G�ƻs���L�a��
'-----------------------------------------------------------

Sub AdvancedFilterSample2()

    Range("�ǲ�").AdvancedFilter xlFilterCopy, Range("����d��"), _
        Worksheets("�z�ﵲ�G").Range("A2")

End Sub


'-----------------------------------------------------------
'�٭�M��
'-----------------------------------------------------------

Sub ResetList()
    
    If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData

End Sub



