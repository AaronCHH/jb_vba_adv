Attribute VB_Name = "Module2"
Option Explicit
  
'------------------------------
'範例96
'控制記事本
'------------------------------

Sub ControlNotePad()
    Dim myPath As String
    Dim myID As Double             'Shell函數的傳回值

    myPath = ActiveWorkbook.Path & "\"

    '開啟記事本
    myID = Shell("Notepad.exe", vbNormalFocus)
    
'<若無法開啟記事本，請加上記事本程式的路徑>
    'myID = Shell("C:\Windows\Notepad.exe", vbNormalFocus)

    '由[Alt]+[F]+[O]鍵傳送開啟[開啟舊檔]對話方塊的命令
    SendKeys "%FO", True

    '指定並傳送開啟檔案名稱的按鍵代碼
    SendKeys myPath & "Report.txt", True

    '由[ENTER]鍵傳送開啟檔案的命令
    SendKeys "{ENTER}", True

    '複製儲存格的內容
    Worksheets("書籍販售量").Range("書籍販售量預估").Copy
    
    '開啟記事本
    AppActivate myID
    
    '由[Ctrl]+{V]鍵傳送貼上複製內容的命令
    SendKeys "^V", True
    
    '解除複製模式
    Application.CutCopyMode = False
End Sub




