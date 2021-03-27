Attribute VB_Name = "Module1"
Option Explicit

'----------------------------------
'�d��71
'�}�ҥH�r�����j����r�ɮ�
'----------------------------------

Sub OpenTxtFile()
    ChDrive ActiveWorkbook.Path
    ChDir ActiveWorkbook.Path
    
    Workbooks.OpenText FileName:="Fuji.txt", _
        DataType:=xlDelimited, Comma:=True
        
    MsgBox "�}�ҡuFuji.txt�v�C" & Chr(13) & _
        "�Y�������A�а��楨���uCloseTxtFile�v�C"
End Sub

Sub CloseTxtFile()
    Workbooks("Fuji.txt").Close False
End Sub

'-----------------------------------------------
'�d��72
'�}�ҥH�r��+�Ŷ����j���ɮ�
'-----------------------------------------------

Sub OpenTxtFile2()
    ChDrive ActiveWorkbook.Path
    ChDir ActiveWorkbook.Path
    
    Workbooks.OpenText FileName:="Fuji2.txt", _
        DataType:=xlDelimited, ConsecutiveDelimiter:=True, _
        Comma:=True, Space:=True
        
        MsgBox "�}�ҡuFuji2.txt�v�C" & Chr(13) & _
            "�Y�������A�а��楨���uCloseTxtFile2�v�C"
End Sub

Sub CloseTxtFile2()
    Workbooks("Fuji2.txt").Close False
End Sub

'-------------------------------
'�d��73
'�N�ƭȸ���ഫ����r
'-------------------------------

Sub OpenTxtFile3()
    ChDrive ActiveWorkbook.Path
    ChDir ActiveWorkbook.Path
    
    Workbooks.OpenText FileName:="Fuji3.txt", _
        DataType:=xlDelimited, Comma:=True, _
        FieldInfo:=Array(Array(1, 2), Array(2, 1), Array(3, 2), _
            Array(4, 1), Array(5, 1), Array(6, 9), _
            Array(7, 2), Array(8, 2))
    
    MsgBox "�}�ҡuFuji3.txt�v�C" & Chr(13) & _
        "�Y�������A�а��楨���uCloseTxtFile3�v�C"
End Sub

Sub CloseTxtFile3()
    Workbooks("Fuji3.txt").Close False
End Sub
