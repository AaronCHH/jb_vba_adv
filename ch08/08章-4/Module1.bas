Attribute VB_Name = "Module1"
Option Explicit

'--------------------------------------
'範例82
'建立檔案後寫入資料(1)
'--------------------------------------

Sub FSOSample1()
    Dim myFSO As Object, myTS As Object
    
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    Set myTS = myFSO.CreateTextFile("C:\FSOSample1.txt", True)
    
    myTS.WriteLine "本文翻譯日期：" & Date
    
    myTS.Close
End Sub


'--------------------------------------
'範例83
'建立檔案後寫入資料(2)
'--------------------------------------

Sub FSOSample2()
    Dim myFSO As New FileSystemObject
    
    With myFSO.CreateTextFile("C:\FSOSample1.txt", True)
    
        .WriteLine "本文翻譯日期：" & Date
    
        .Close
    End With
End Sub


'--------------------------------------
'範例84
'檢查磁碟的容量
'--------------------------------------

Sub FSOSample3()
    Dim myFSO As New FileSystemObject
    Dim myDS1 As Variant, myDS2 As Variant
    
    With myFSO.GetDrive("A")
        
        myDS1 = .TotalSize
        
        myDS2 = .AvailableSpace
    
    End With
    
    MsgBox "已使用空間：" & Format(myDS1 - myDS2, "#,##0") & vbCrLf & _
        "可用空間：" & Format(myDS2, "#,##0") & vbCrLf & vbCrLf & _
        "容量：" & Format(myDS1, "#,##0") & vbCrLf
End Sub


'--------------------------------------
'範例85
'查詢磁碟的類型
'--------------------------------------

Sub FSOSample4()
    Dim myFSO As New FileSystemObject
    Dim myDrv As Drive
    Dim myMsg As String
    
    For Each myDrv In myFSO.Drives
        
        myMsg = myMsg & myDrv.DriveLetter & "："
        
        Select Case myDrv.DriveType
            Case 0
                myMsg = myMsg & "不明" & vbCrLf
            Case 1
                myMsg = myMsg & "抽取式磁碟" & vbCrLf
            Case 2
                myMsg = myMsg & "硬式磁碟" & vbCrLf
            Case 3
                myMsg = myMsg & "網路磁碟" & vbCrLf
            Case 4
                myMsg = myMsg & "CD-ROM" & vbCrLf
            Case 5
                myMsg = myMsg & "RAM Disk" & vbCrLf
        End Select
    Next
    
    MsgBox myMsg
End Sub


'--------------------------------------
'範例86
'檢查磁碟機的準備狀態
'--------------------------------------

Sub FSOSample5()
    Dim myFSO As New FileSystemObject
    
    If myFSO.Drives("A").IsReady = True Then
        FileCopy "C:\Excel2003VBA應用篇\Fuji.txt", "A:\Fuji.txt"
    Else
        MsgBox "未插入磁碟片"
    End If
End Sub


'--------------------------------------
'範例87
'取得子資料夾
'--------------------------------------

Sub FSOSample6()
    Dim myFSO As New FileSystemObject
    Dim myFld As Folder
    Dim i As Integer
    
    Worksheets("Sheet2").Activate
    i = 1
    
    With myFSO.GetFolder("C:\WINNT")
    
    '若發生錯誤，請將上行陳述式加上註記，然後執行下行陳述式
    'With myFSO.GetFolder("C:\Windows")
    
        For Each myFld In .SubFolders
        
            i = i + 1
            Cells(i, 1).Value = myFld.Name
            
        Next
    End With
End Sub


'------------------------------------------------------
'範例88
'取得資料夾及其子資料夾內的檔案總容量
'------------------------------------------------------

Sub FSOSample7()
    Dim myFSO As New FileSystemObject
    Dim mySize1 As Variant, mySize2 As Variant
    
    With myFSO.GetFolder("C:\My Documents")
    
        mySize1 = .Size
        mySize2 = mySize1 / 1024 / 1024
        
        MsgBox _
            "C:\My Documents資料夾全部的檔案大小為：" & vbCrLf & vbCrLf & _
            Format(mySize2, "#,##0.0") & "MB" & " (" & _
            Format(mySize1, "#,##0") & "Byte)"
    
    End With
End Sub




