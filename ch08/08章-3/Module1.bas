Attribute VB_Name = "Module1"
Option Explicit

'----------------------------
'�d��79
'�R����Ƨ������ɮ�

Sub KillFile()
    Dim myPath As String
    
    myPath = ActiveWorkbook.Path & "\"
    
    If Dir(myPath & "DataBook.xls") <> "" Then
        Kill myPath & "DataBook.xls"
        
        MsgBox "DataBook.xls�R�������A" & Chr(13) & _
            "�Y���A���إ�DataBook.xls�A�а��楨���uMakeDataBook�v�C"
    Else
        MsgBox "�䤣��DataBook.xls"
    End If
End Sub

'�إ�DataBook.xls
Sub MakeDataBook()
    Dim myPath As String
    
    myPath = ActiveWorkbook.Path & "\"
    
    FileCopy myPath & "Dummy.xls", myPath & "DataBook.xls"
        
    MsgBox "DataBook.xls�إߧ����C"
End Sub


'--------------------------------
'�d��80
'�j�M��Ƨ����ŦX���󪺽Ƽ��ɮ�
'--------------------------------

Sub SearchFile()
    Dim myPath As String
    Dim myFName As String
    Dim i As Integer
    
    Worksheets("�j�M�ɮ�").Activate
    i = 1
    Cells(i, 1).Value = "�ɮצW��"
    Cells(i, 2).Value = "�ɮפj�p"
    Cells(i, 3).Value = "�ɮ׫إ�/�ק���"
    
    myPath = ActiveWorkbook.Path & "\"
    
    myFName = Dir(myPath & "*.xls")
    
    Do While myFName <> ""
        i = i + 1
        Cells(i, 1).Value = myFName
        Cells(i, 2).Value = FileLen(myPath & myFName)
        Cells(i, 3).Value = FileDateTime(myPath & myFName)
        
        myFName = Dir()
    Loop
End Sub

'-------------------------------
'�d��81
'FileSearch���B��
'-------------------------------

Sub UseFileSearch()
    Dim myFSObj As FileSearch
    Dim i As Integer
    
    MsgBox "�̧ǦC�X�P8��-3.xls��b�P��Ƨ���Excel����ï�W�١C"
    
    Worksheets("�j�M�ɮ�").Activate
            
    Set myFSObj = Application.FileSearch
    
    With myFSObj
        .LookIn = ActiveWorkbook.Path
        .Filename = "*.xls"
        
        If .Execute(SortBy:=msoSortByFileName, _
            SortOrder:=msoSortOrderAscending) > 0 Then
            
            MsgBox "�j�M�쪺Excel����ï�ƶq���G" & .FoundFiles.Count & " �ӡC"
            
            For i = 1 To .FoundFiles.Count
                Cells(i, 1).Value = .FoundFiles(i)
            Next i
        
        Else
            MsgBox "�䤣��Excel����ï�C"
        End If
    End With
End Sub
