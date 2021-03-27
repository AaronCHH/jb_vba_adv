Attribute VB_Name = "Module1"
Option Explicit
    Dim objWord As New Word.Application     'Word應用程式
    Dim objWordDoc As Word.Document         '新增的Word文件

'-----------------------------------------------------------
'範例93
'使用早期繫結控制Word物件
'-----------------------------------------------------------

Sub CreateWordApp()
    Dim myPath As String
    Dim myReturn As Integer
    
    myPath = ActiveWorkbook.Path & "\"
    
    myReturn = MsgBox(prompt:="「Test.doc」是關閉的嗎？" & _
        "(若是開啟的，將發生錯誤)", Buttons:=vbYesNo)
    If myReturn = vbNo Then Exit Sub
    
    With objWord
        .Visible = True                                               '顯示Word
        .WindowState = wdWindowStateMaximize '將視窗最大化
        .Documents.Add                                            '新增文件
        
        '將新增文件代入物件變數中
        Set objWordDoc = .ActiveDocument
    End With
    
    '在文件中插入文字
    With objWord.Selection
        .InsertAfter "建立Word物件的測試"
        .InsertParagraphAfter
        .InsertAfter Now() & " 建立"
        .MoveRight
    End With
    
    '設定段落1的格式
    With objWordDoc.Paragraphs(1).Range
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Font
            .Name = "標楷體"
            .Size = 20
            .Bold = True
        End With
    End With
    
    '設定段落2的格式
    objWordDoc.Paragraphs(2).Range.ParagraphFormat _
        .Alignment = wdAlignParagraphRight
    
    Application.Wait Now() + TimeValue("00:00:03")
    
    objWord.WindowState = wdWindowStateMinimize '將視窗最小化

    MsgBox "啟動Word並建立新文件" & Chr(13) & _
        "按下[確定]按鈕後儲存文件並關閉Word"

    objWordDoc.SaveAs myPath & "Test.doc"   '將新增的文件另存新檔
    objWordDoc.Close                        '關閉新增的文件
    objWord.Quit                            '關閉Word

    Set objWord = Nothing                   '清除物件變數的內容
    Set objWordDoc = Nothing
End Sub


'-----------------------------------------------------------
'範例94
'將Excel工作表傳送到Word文件中
'-----------------------------------------------------------

Sub CreateWordApp2()
    
    With objWord
        .Visible = True                      '顯示Word
        .WindowState = wdWindowStateMaximize '將視窗最大化
        .Documents.Open ActiveWorkbook.Path & "\Report.doc" '開啟Report.doc
        
        '將Report.doc代入物件變數中
        Set objWordDoc = .ActiveDocument
        
        '在文件中插入文字
        With .Selection
            .Move Count:=objWordDoc.Characters.Count
            .InsertParagraphAfter
            .InsertAfter "CD販售張數"
            .InsertParagraphAfter
            .MoveRight
        End With
    End With
    
    '複製儲存格的資料
    Worksheets("CD販售量").Range("CD販賣張數").Copy
    
    '貼到Word內
    With objWord.Selection
        .Paste
        .TypeParagraph
    End With
    
    '複製圖表
    Worksheets("CD販售量").ChartObjects(1).Copy
    
    '設定貼到Word的格式
    With objWord
        .Selection.PasteSpecial Placement:=wdInLine, _
            DataType:=wdPasteMetafilePicture
        .Selection.ParagraphFormat.Alignment = _
            wdAlignParagraphCenter
    End With
    
    '列印(列印時將中斷巨集的執行)
    objWord.PrintOut Background:=False
    
    '不儲存文件，關閉Word
    objWordDoc.Close SaveChanges:=False
    
    objWord.Quit                '關閉Word
    
    Set objWord = Nothing       '清除物件變數的內容
    Set objWordDoc = Nothing
End Sub


'----------------------------------
'範例95
'存取已啟動的Word
'----------------------------------

Sub GetWordApp()
    On Error GoTo HandleErr
    Dim myAppOpen As Boolean
    
    Set objWord = GetObject(, "Word.Application")
    myAppOpen = True
    
MacroContinue:
    If myAppOpen = False Then
        Set objWord = CreateObject("Word.Application")
    End If
    
    With objWord
        .Visible = True
        .WindowState = wdWindowStateMinimize
        .Documents.Add
    End With
    
    Set objWord = Nothing
    
    Exit Sub

HandleErr:
    '若尚未將ActiveX元件建立為物件
    If Err.Number = 429 Then
        myAppOpen = False
        Resume MacroContinue
    End If
End Sub


