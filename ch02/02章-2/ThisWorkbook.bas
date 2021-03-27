Option Explicit

'-----------------------------------------------------
'範例14
'新增工作表時執行事件程序
'-----------------------------------------------------

Private Sub Workbook_NewSheet(ByVal Sh As Object)

    MsgBox "將新增的" & Sh.Name & "移到活頁簿最後面。"
    Sh.Move After:=Sheets(Sheets.Count)

End Sub

'---------------------------------------------------
'範例15
'關閉活頁簿時執行事件程序
'---------------------------------------------------

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    If Sheet2.Range("B1").Value = "" Then
        MsgBox "關閉活頁簿時" & vbCrLf & _
            "請在Sheet2的B1儲存格中輸入建立者。"
        Sheet2.Activate
        Range("B1").Activate
    
        Cancel = True
    End If

End Sub