Attribute VB_Name = "Module1"
Option Explicit

'------------------------------------
'範例21
'On Error GoTo陳述式範例
'------------------------------------

Sub TrapSample()
    
    '啟用Error Trap功能
    On Error GoTo HandleErr
    
    '預測會發生錯誤的陳述式
    ActiveWorkbook.Charts(1).Activate
    
    '沒有發生錯誤的處理程序
    ActiveChart.SizeWithWindow = True
    MsgBox "執行完成，沒有發生錯誤。"
    Exit Sub
    
HandleErr:
    MsgBox "圖表工作表不存在。"
End Sub

'-------------------------------------------
'範例22
'On Error Resume Next陳述式範例
'-------------------------------------------

Sub TrapSample2()
    Dim myRange As Range
    Dim myPrompt As String, myTitle As String
    
    Worksheets("Sheet3").Activate
    Cells.Clear
    
    myPrompt = "被選取的儲存格範圍將輸入「ABC」" & vbCr & _
        "請使用滑鼠選取儲存格範圍。"
    myTitle = "輸入儲存格範圍"
    
    '啟用Error Trap功能
    On Error Resume Next
    
    '預測會發生錯誤的陳述式
    Set myRange = Application.InputBox(Prompt:=myPrompt, _
        Title:=myTitle, Type:=8)
    
    '判斷前面的陳述式是否發生錯誤
    If myRange Is Nothing Then Exit Sub

    myRange.Value = "ABC"
End Sub

'-----------------------------
'範例23
'查詢錯誤代碼及錯誤訊息
'-----------------------------

Sub DisplayErr()
    Dim myMsg As String
    
    Worksheets("Sheet4").Activate
    
    On Error GoTo HandleErr
    
    Range("B3").Value = Range("B1").Value / Range("B2").Value
    
    Exit Sub

HandleErr:
    myMsg = "錯誤代碼：" & Err.Number & vbCrLf & _
        "錯誤訊息：" & Err.Description
    MsgBox myMsg

    Range("B3").Value = 0
End Sub

'--------------------------------
'範例24
'依據錯誤種類判斷條件
'--------------------------------

Sub OpenFile()
    Dim myFD As Variant, myFN As Variant
    Dim myPrompt As String, myMsg As String
    Dim myBuf As String
    
    MsgBox "請在磁碟機的根目錄下準備一個文字檔案，" _
        & vbCr & "不限制檔案名稱。"
    
InputFD:
    myPrompt = "請輸入磁碟機代號："
    myFD = Application.InputBox(Prompt:=myPrompt, Default:="A")
    If VarType(myFD) <> vbString Then Exit Sub
    
InputFN:
    myPrompt = "請輸入檔案名稱："
    myFN = Application.InputBox(Prompt:=myPrompt)
    If VarType(myFN) <> vbString Then Exit Sub
    
    On Error GoTo HandleErr
    
    Open myFD & ":\" & myFN For Input As #1
    
    Do Until EOF(1)
        Line Input #1, myBuf
    Loop
    
    MsgBox "處理完成，沒有發生錯誤。"
    Close #1
    
    Exit Sub

HandleErr:
    Select Case Err.Number
        Case 53                 '找不到檔案
            MsgBox Err.Description & vbCr & _
                 "請重新輸入檔案名稱："
            Resume InputFN
            
        Case 55                 '檔案已開啟
            MsgBox Err.Description
            Resume Next
        
        Case 68, 75, 76         '週邊設備無法使用
            MsgBox Err.Description & vbCr & vbCr & _
                "指定的磁碟無效，" & vbCr & _
                "請再輸入磁磁代號："
            Resume InputFD
        
        Case 52, 71             '磁碟尚未就緒
            myMsg = Err.Description & vbCr & _
                "要插入磁片繼續嗎？"
            If MsgBox(myMsg, vbExclamation + vbYesNo) = vbYes Then
                Resume
            Else
                Exit Sub
            End If
    End Select
End Sub

'------------------------------------
'範例25
'將值輸出到即時運算視窗
'------------------------------------

Sub OutToWindow()
    Dim n As Integer, m As Integer
    
    For n = 1 To 10
        m = 2 ^ n
        Debug.Print "m=" & m        '開啟即時運算視窗
    Next
    
    'Stop
End Sub

