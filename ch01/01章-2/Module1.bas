Attribute VB_Name = "Module1"
Option Explicit

'---------------------------------------------
'範例3
'使用引用呼叫其他活頁簿的程序
'---------------------------------------------

Sub SansyoSettei()
    Call Sample1
End Sub

'------------------------------------------------
'範例4
'使用Run方法呼叫其他活頁簿的程序
'(不指定路徑)
'------------------------------------------------

Sub AppRun()
    Application.Run "Subrtin2.xls!Sample2"
End Sub

'-------------------------------------------------
'範例5
'使用指定的路徑呼叫其他活頁簿的程序
'-------------------------------------------------

Sub AppRun2()
    Dim myWBPath As String
    
    myWBPath = ActiveWorkbook.Path
    
    Application.Run "'" & myWBPath & "\Subrtin2.xls'!Sample2"
End Sub

