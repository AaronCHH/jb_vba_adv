Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

'----------------------------------------
'範例75
'不開啟活頁簿而能讀入文字檔案
'----------------------------------------

Sub ReadTxt()
    Dim myTxtFile As String
    Dim myBuf(6) As String
    Dim i As Integer, j As Integer
    
    Application.ScreenUpdating = False
    
    myTxtFile = ActiveWorkbook.Path & "\Fuji.txt"
    
    Worksheets("郵遞區號").Activate
    
    Open myTxtFile For Input As #1
    
    Do Until EOF(1)
        Input #1, myBuf(1), myBuf(2), myBuf(3), myBuf(4), _
            myBuf(5), myBuf(6)
    
    '將資料置入儲存格內
        i = i + 1
        For j = 1 To 6
            Cells(i, j) = myBuf(j)
        Next j
    Loop
    
    Close #1
End Sub


'----------------------------------
'範例76
'輸出文件類型的文字檔案
'----------------------------------

Sub ReadTxt2()
    Dim myTxtFile As String, myFNo As Integer, myBuf As String
    Dim i As Integer
       
    Application.ScreenUpdating = False
    
    myTxtFile = ActiveWorkbook.Path & "\WordPro.txt"
    
    Worksheets("文件類型").Activate
        
    myFNo = FreeFile            '取得可使用的檔案代碼
    Open myTxtFile For Input As #myFNo
    
    Do Until EOF(myFNo)
        Line Input #myFNo, myBuf
    
        i = i + 1
        Cells(i, 1) = myBuf    '將資料置入儲存格內
    Loop
    
    Close #myFNo
End Sub


'-----------------------------------
'範例77
'將工作表的內容儲存為CSV格式檔案
'-----------------------------------

Sub WriteCsv()
    Dim myTxtFile As String, myFNo As Integer
    Dim myLastRow As Long, i As Long
       
    Application.ScreenUpdating = False
    
    myTxtFile = ActiveWorkbook.Path & "\Numazu.csv"
    
    Worksheets("郵遞區號2").Activate
    myLastRow = Range("A1").CurrentRegion.Rows.Count
    
    myFNo = FreeFile
    Open myTxtFile For Output As #myFNo
    
    For i = 1 To myLastRow
        Write #myFNo, Cells(i, 1), Cells(i, 2), Cells(i, 3), _
            Cells(i, 4), Cells(i, 5), Cells(i, 6)
    Next
    
    Close #myFNo
                
    MsgBox "「郵遞區號2」工作表內的資料將建立為「Numazu.csv」檔案。"
End Sub


'-----------------------------------
'範例78
'將工作表的內容儲存為文件類型檔案
'-----------------------------------

Sub WriteTxt()
    Dim myTxtFile As String, myFNo As Integer
    Dim myLastRow As Long, i As Long
       
    Worksheets("文件類型2").Activate
    Application.ScreenUpdating = False
    
    myTxtFile = ActiveWorkbook.Path & "\Column.txt"
    
    myLastRow = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    
    myFNo = FreeFile
    Open myTxtFile For Output As #myFNo
    
    For i = 1 To myLastRow
        Print #myFNo, Cells(i, 1)
    Next
    
    Close #myFNo
                
    MsgBox "此工作表的資料將建立為為「Column.txt」檔案。"
End Sub
