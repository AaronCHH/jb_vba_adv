Attribute VB_Name = "Module3"
Option Explicit

'-------------------------------------------------------
'範例97
'關閉以Shell函數啟動的應用程式之前進入待機狀態
'-------------------------------------------------------

'傳回既存物件控制代碼的函數
'傳回值　成功 = 指定處理的Open控制代碼
' 　　　　失敗 = NULL
Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Public Const PROCESS_QUERY_INFORMATION = &H400&


'傳回指定處理結束狀態的函數
'傳回值　成功 = 0以外的數值
'　　　　失敗 = 0
Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, _
    lpExitCode As Long) As Long
        
'判斷指定的處理是否結束
'(若尚未結束，將置入STILL_ACTIVE)
Public Const STATUS_PENDING = &H103&
Public Const STILL_ACTIVE = STATUS_PENDING


'關閉開啟中物件控制代碼的函數
''傳回值　成功 = 0以外的數值
'　　　　失敗 = 0
Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long


'程序
Sub GetExitCodeProcess_Sample()
    Dim lngProcessId As Long    'Shell函數的傳回值
    Dim lngProcess As Long      'OpenProcess函數的傳回值
    Dim lngExitCode As Long     '結束程式碼
    Dim rc As Long
     
    MsgBox "按[確定]按鈕後將開啟記事本，關閉記事本之前程序將處於暫停的狀態"
              
    '啟動記事本
    lngProcessId = Shell("Notepad.exe", vbNormalFocus)
    
    '取得以Shell函數啟動的應用程式的處理物件的控制代碼
    lngProcess = OpenProcess(PROCESS_QUERY_INFORMATION, _
                            1, _
                            lngProcessId)
                            
    '以GetExitCodeProcess取得處理的結束狀態
    '若啟動的應用程式處於尚未關閉的狀態，將由DoEvent繼續對作業系統詢問其狀態
    Do
        rc = GetExitCodeProcess(lngProcess, lngExitCode)
        DoEvents
    Loop While lngExitCode = STILL_ACTIVE

    '關閉開啟中的物件控制代碼
    rc = CloseHandle(lngProcess)

    MsgBox "記事本關閉"
End Sub



