Attribute VB_Name = "Module2"
Option Explicit

'------------------------------------------------------
'範例74
'開啟固定長度欄位的文字檔案
'------------------------------------------------------

Sub OpenTxtFile4()
    ChDrive ActiveWorkbook.Path
    ChDir ActiveWorkbook.Path
    
    Workbooks.OpenText FileName:="Nyukin.txt", DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 2), Array(5, 5), Array(13, 1), _
            Array(17, 1), Array(36, 1))
    
    MsgBox "開啟「Nyukin.txt」。" & Chr(13) & _
        "若欲關閉，請執行巨集「CloseTxtFile4」。"
End Sub

Sub CloseTxtFile4()
    Workbooks("Nyukin.txt").Close False
End Sub
