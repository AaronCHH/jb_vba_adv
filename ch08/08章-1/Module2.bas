Attribute VB_Name = "Module2"
Option Explicit

'------------------------------------------------------
'�d��74
'�}�ҩT�w������쪺��r�ɮ�
'------------------------------------------------------

Sub OpenTxtFile4()
    ChDrive ActiveWorkbook.Path
    ChDir ActiveWorkbook.Path
    
    Workbooks.OpenText FileName:="Nyukin.txt", DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 2), Array(5, 5), Array(13, 1), _
            Array(17, 1), Array(36, 1))
    
    MsgBox "�}�ҡuNyukin.txt�v�C" & Chr(13) & _
        "�Y�������A�а��楨���uCloseTxtFile4�v�C"
End Sub

Sub CloseTxtFile4()
    Workbooks("Nyukin.txt").Close False
End Sub
