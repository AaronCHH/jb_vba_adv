Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

'----------------------------------------
'�d��75
'���}�Ҭ���ï�ӯ�Ū�J��r�ɮ�
'----------------------------------------

Sub ReadTxt()
    Dim myTxtFile As String
    Dim myBuf(6) As String
    Dim i As Integer, j As Integer
    
    Application.ScreenUpdating = False
    
    myTxtFile = ActiveWorkbook.Path & "\Fuji.txt"
    
    Worksheets("�l���ϸ�").Activate
    
    Open myTxtFile For Input As #1
    
    Do Until EOF(1)
        Input #1, myBuf(1), myBuf(2), myBuf(3), myBuf(4), _
            myBuf(5), myBuf(6)
    
    '�N��Ƹm�J�x�s�椺
        i = i + 1
        For j = 1 To 6
            Cells(i, j) = myBuf(j)
        Next j
    Loop
    
    Close #1
End Sub


'----------------------------------
'�d��76
'��X�����������r�ɮ�
'----------------------------------

Sub ReadTxt2()
    Dim myTxtFile As String, myFNo As Integer, myBuf As String
    Dim i As Integer
       
    Application.ScreenUpdating = False
    
    myTxtFile = ActiveWorkbook.Path & "\WordPro.txt"
    
    Worksheets("�������").Activate
        
    myFNo = FreeFile            '���o�i�ϥΪ��ɮץN�X
    Open myTxtFile For Input As #myFNo
    
    Do Until EOF(myFNo)
        Line Input #myFNo, myBuf
    
        i = i + 1
        Cells(i, 1) = myBuf    '�N��Ƹm�J�x�s�椺
    Loop
    
    Close #myFNo
End Sub


'-----------------------------------
'�d��77
'�N�u�@�����e�x�s��CSV�榡�ɮ�
'-----------------------------------

Sub WriteCsv()
    Dim myTxtFile As String, myFNo As Integer
    Dim myLastRow As Long, i As Long
       
    Application.ScreenUpdating = False
    
    myTxtFile = ActiveWorkbook.Path & "\Numazu.csv"
    
    Worksheets("�l���ϸ�2").Activate
    myLastRow = Range("A1").CurrentRegion.Rows.Count
    
    myFNo = FreeFile
    Open myTxtFile For Output As #myFNo
    
    For i = 1 To myLastRow
        Write #myFNo, Cells(i, 1), Cells(i, 2), Cells(i, 3), _
            Cells(i, 4), Cells(i, 5), Cells(i, 6)
    Next
    
    Close #myFNo
                
    MsgBox "�u�l���ϸ�2�v�u�@������ƱN�إ߬��uNumazu.csv�v�ɮסC"
End Sub


'-----------------------------------
'�d��78
'�N�u�@�����e�x�s����������ɮ�
'-----------------------------------

Sub WriteTxt()
    Dim myTxtFile As String, myFNo As Integer
    Dim myLastRow As Long, i As Long
       
    Worksheets("�������2").Activate
    Application.ScreenUpdating = False
    
    myTxtFile = ActiveWorkbook.Path & "\Column.txt"
    
    myLastRow = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    
    myFNo = FreeFile
    Open myTxtFile For Output As #myFNo
    
    For i = 1 To myLastRow
        Print #myFNo, Cells(i, 1)
    Next
    
    Close #myFNo
                
    MsgBox "���u�@����ƱN�إ߬����uColumn.txt�v�ɮסC"
End Sub
