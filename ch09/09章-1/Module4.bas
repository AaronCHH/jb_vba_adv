Attribute VB_Name = "Module4"
Option Explicit

'-------------------------------------------------------
'範例98
'避免重複啟動應用程式
'-------------------------------------------------------


'從Class名稱或Caption取得視窗控制代碼的函數
'傳回值　成功 = 視窗控制代碼
' 　　　　失敗 = NULL
Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" _
    (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

'程序
Sub FindWindow_Sample()
    Dim strClassName As String  'Class名稱
    Dim rc As Long
    
    Dim lngProcessId As Long    'Shell函數的傳回值
    
    '指定Class名稱
    strClassName = "SciCalc"
        
    '取得小算盤的視窗控制代碼
    rc = FindWindow(strClassName, _
                    vbNullString)
                    
    '取得視窗控制代碼但尚未啟動
    If rc <> 0& Then
        MsgBox "小算盤已開啟"
        Exit Sub
    End If
        
    '啟動小算盤
    lngProcessId = Shell("Calc.exe", vbNormalFocus)
End Sub




