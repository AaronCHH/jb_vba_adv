Attribute VB_Name = "Module1"
Option Explicit

'----------------------------
'範例79
'刪除資料夾內的檔案

Sub KillFile()
    Dim myPath As String
    
    myPath = ActiveWorkbook.Path & "\"
    
    If Dir(myPath & "DataBook.xls") <> "" Then
        Kill myPath & "DataBook.xls"
        
        MsgBox "DataBook.xls刪除完成，" & Chr(13) & _
            "若欲再次建立DataBook.xls，請執行巨集「MakeDataBook」。"
    Else
        MsgBox "找不到DataBook.xls"
    End If
End Sub

'建立DataBook.xls
Sub MakeDataBook()
    Dim myPath As String
    
    myPath = ActiveWorkbook.Path & "\"
    
    FileCopy myPath & "Dummy.xls", myPath & "DataBook.xls"
        
    MsgBox "DataBook.xls建立完成。"
End Sub


'--------------------------------
'範例80
'搜尋資料夾內符合條件的複數檔案
'--------------------------------

Sub SearchFile()
    Dim myPath As String
    Dim myFName As String
    Dim i As Integer
    
    Worksheets("搜尋檔案").Activate
    i = 1
    Cells(i, 1).Value = "檔案名稱"
    Cells(i, 2).Value = "檔案大小"
    Cells(i, 3).Value = "檔案建立/修改日期"
    
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
'範例81
'FileSearch的運用
'-------------------------------

Sub UseFileSearch()
    Dim myFSObj As FileSearch
    Dim i As Integer
    
    MsgBox "依序列出與8章-3.xls位在同資料夾的Excel活頁簿名稱。"
    
    Worksheets("搜尋檔案").Activate
            
    Set myFSObj = Application.FileSearch
    
    With myFSObj
        .LookIn = ActiveWorkbook.Path
        .Filename = "*.xls"
        
        If .Execute(SortBy:=msoSortByFileName, _
            SortOrder:=msoSortOrderAscending) > 0 Then
            
            MsgBox "搜尋到的Excel活頁簿數量為：" & .FoundFiles.Count & " 個。"
            
            For i = 1 To .FoundFiles.Count
                Cells(i, 1).Value = .FoundFiles(i)
            Next i
        
        Else
            MsgBox "找不到Excel活頁簿。"
        End If
    End With
End Sub
