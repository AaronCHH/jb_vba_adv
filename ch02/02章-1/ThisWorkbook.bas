Option Explicit

'---------------------------------------------
'範例13
'開啟活頁簿時執行事件程序
'---------------------------------------------

Private Sub Workbook_Open()
    MsgBox "將同時開啟Dummy.xls"
    'Workbooks.Open FileName:="Dummy.xlsx"
    Workbooks.Open FileName:="test.xlsx"
End Sub

