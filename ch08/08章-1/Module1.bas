Attribute VB_Name = "Module1"
Option Explicit

'----------------------------------
'範例71
'開啟以逗號分隔的文字檔案
'----------------------------------

Sub OpenTxtFile()
    ChDrive ActiveWorkbook.Path
    ChDir ActiveWorkbook.Path
    
    Workbooks.OpenText FileName:="Fuji.txt", _
        DataType:=xlDelimited, Comma:=True
        
    MsgBox "開啟「Fuji.txt」。" & Chr(13) & _
        "若欲關閉，請執行巨集「CloseTxtFile」。"
End Sub

Sub CloseTxtFile()
    Workbooks("Fuji.txt").Close False
End Sub

'-----------------------------------------------
'範例72
'開啟以逗號+空間分隔的檔案
'-----------------------------------------------

Sub OpenTxtFile2()
    ChDrive ActiveWorkbook.Path
    ChDir ActiveWorkbook.Path
    
    Workbooks.OpenText FileName:="Fuji2.txt", _
        DataType:=xlDelimited, ConsecutiveDelimiter:=True, _
        Comma:=True, Space:=True
        
        MsgBox "開啟「Fuji2.txt」。" & Chr(13) & _
            "若欲關閉，請執行巨集「CloseTxtFile2」。"
End Sub

Sub CloseTxtFile2()
    Workbooks("Fuji2.txt").Close False
End Sub

'-------------------------------
'範例73
'將數值資料轉換為文字
'-------------------------------

Sub OpenTxtFile3()
    ChDrive ActiveWorkbook.Path
    ChDir ActiveWorkbook.Path
    
    Workbooks.OpenText FileName:="Fuji3.txt", _
        DataType:=xlDelimited, Comma:=True, _
        FieldInfo:=Array(Array(1, 2), Array(2, 1), Array(3, 2), _
            Array(4, 1), Array(5, 1), Array(6, 9), _
            Array(7, 2), Array(8, 2))
    
    MsgBox "開啟「Fuji3.txt」。" & Chr(13) & _
        "若欲關閉，請執行巨集「CloseTxtFile3」。"
End Sub

Sub CloseTxtFile3()
    Workbooks("Fuji3.txt").Close False
End Sub
