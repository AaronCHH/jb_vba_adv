Attribute VB_Name = "Module1"
Option Explicit

'-----------------------------------
'範例36
'變更Excel標題列的文字
'-----------------------------------

Sub ChangeTitleBar()
    Application.Caption = "旗標出版(股)公司"
    ActiveWindow.Caption = "文件排版系統"
End Sub


'----------------------------------------
'範例37
'將視窗貼齊窗格
'----------------------------------------

Sub MaximizeWindow()
    With ActiveWindow
        .WindowState = xlNormal
        .Top = 0
        .Left = 0
        .Height = Application.UsableHeight
        .Width = Application.UsableWidth
    End With
End Sub


'----------------------------------------
'範例38
'將Excel放到螢幕顯示區域之外
'----------------------------------------

Sub HideExcel()
    Dim myTop As Double, myLeft As Double
    
    Application.WindowState = xlNormal
    
    myTop = Application.Top
    myLeft = Application.Left
    
    MsgBox "隱藏Excel"
    Application.Left = -Application.Width
    
    MsgBox "再度顯示Excel"
    Application.Top = myTop
    Application.Left = myLeft
End Sub


'----------------------------------------
'範例39
'關閉螢幕更新功能
'----------------------------------------

Sub StopScreen()
    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    For i = 1 To 10
        Worksheets(1).Activate
        Worksheets(2).Activate
    Next i
    
    Application.ScreenUpdating = True
End Sub


'----------------------------------------
'範例40
'隱藏刪除工作表的確認訊息
'----------------------------------------

Sub StopAlertMsg()
    Application.DisplayAlerts = False

    Worksheets.Add.Name = "Dummy"
    MsgBox "新增工作表「Dummy」" & vbCrLf & _
        "接著，刪除「Dummy」"

    Worksheets("Dummy").Delete
    
    Application.DisplayAlerts = True
End Sub


'----------------------------------------
'範例41
'在狀態列顯示訊息
'----------------------------------------

Sub DispStatusBar()
    Dim myStatusBar As Boolean
    Dim myCell As Range
    
    Worksheets("Sheet3").Activate
    
    myStatusBar = Application.DisplayStatusBar
    
    Application.DisplayStatusBar = True
    
    For Each myCell In Range("A1:C5")
        myCell.Value = "ABC"
        
        Application.StatusBar = "寫入 " & myCell.Address & " 中"

        Application.Wait Now + TimeValue("00:00:01")
    Next myCell
   
    Application.StatusBar = False
        
    Application.DisplayStatusBar = myStatusBar
End Sub


'----------------------------------------
'範例42
'內建對話方塊的條件判斷
'----------------------------------------

Sub DispBuiltinDialog()
    Dim myRtn As Boolean
    
    myRtn = Application.Dialogs(xlDialogOptionsView).Show
    
    If myRtn = False Then
        MsgBox "點選「取消」" & vbCrLf & _
            "處理完畢"
        Exit Sub
    End If

    MsgBox "繼續進行處理"
'
' 進行處理
'
End Sub


'----------------------------------------------
'範例43
'使用GetOpenFilename方法開啟文字檔案
'----------------------------------------------

Sub OpenTxtFile()
    Dim myFName As String
    
    myFName = Application. _
        GetOpenFilename("文字檔案(*.prn; *.txt; *.csv),*.prn;*.txt;*.csv")
    
    If myFName <> "False" Then
        Workbooks.OpenText Filename:=myFName, Comma:=True
    End If
End Sub

