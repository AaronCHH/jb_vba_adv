Attribute VB_Name = "Module1"
'-----------------------------------------
'�d��1
'�q�D�{�ǩI�s�Ƶ{��
'(�Ƨǡ��z����C�L�����Φ۰ʿz��)
'-----------------------------------------

Sub PrintMember()
    SortData        '�Ƶ{��(�Ƨ�)
    
    SelectData      '�Ƶ{��(�z��)
    
    PrintData       '�Ƶ{��(�C�L)
    
    KaijoFilter     '�Ƶ{��(���Φ۰ʿz��)
End Sub

'-----------
'�Ƶ{��
'-----------

'1.�Ƨ�
Sub SortData()
    Range("A1").Sort Key1:=Range("D3"), Order1:=xlAscending, _
        Key2:=Range("F3"), Order2:=xlAscending, Header:=xlGuess
End Sub

'2.�z��
Sub SelectData()
    Range("A1").AutoFilter Field:=9, Criteria1:=">=2003/1/1", _
        Operator:=xlAnd, Criteria2:="<=2006/12/31"
End Sub

'3.�C�L���
Sub PrintData()
    Range("A1").CurrentRegion.Offset(1).Select
    Selection.Resize(Selection.Rows.Count - 1).Select
    ActiveSheet.PageSetup.PrintArea = Selection.Address
    ActiveSheet.PrintOut
End Sub

'4.���Φ۰ʿz��
Sub KaijoFilter()
    Range("A1").Select
    Selection.AutoFilter
End Sub
