Attribute VB_Name = "Module1"
Option Explicit


'-----------------------------------------------------------
'�d��110
'�HListRow�����C�i��ާ@
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
'�d��111
'�̧���ܦ���ҦW��
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
'�d��112
'����u���~�W�١v�檺�Ҧ����
'-----------------------------------------------------------

Sub ListColumnSample2()
    Dim myList As ListObject
    
    Set myList = ActiveSheet.ListObjects(1)
    
    myList.ListColumns("���~�W��").Range.Select
End Sub


'-----------------------------------------------------------
'�d��113
'�X�p�C
'-----------------------------------------------------------

Sub Syukeigyo()
    Dim myList As ListObject
    
    Set myList = ActiveSheet.ListObjects(1)
    
    myList.ShowTotals = True
    
    myList.TotalsRowRange.End(xlToRight).Select
End Sub

