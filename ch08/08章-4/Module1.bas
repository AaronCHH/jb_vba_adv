Attribute VB_Name = "Module1"
Option Explicit

'--------------------------------------
'�d��82
'�إ��ɮ׫�g�J���(1)
'--------------------------------------

Sub FSOSample1()
    Dim myFSO As Object, myTS As Object
    
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    Set myTS = myFSO.CreateTextFile("C:\FSOSample1.txt", True)
    
    myTS.WriteLine "����½Ķ����G" & Date
    
    myTS.Close
End Sub


'--------------------------------------
'�d��83
'�إ��ɮ׫�g�J���(2)
'--------------------------------------

Sub FSOSample2()
    Dim myFSO As New FileSystemObject
    
    With myFSO.CreateTextFile("C:\FSOSample1.txt", True)
    
        .WriteLine "����½Ķ����G" & Date
    
        .Close
    End With
End Sub


'--------------------------------------
'�d��84
'�ˬd�ϺЪ��e�q
'--------------------------------------

Sub FSOSample3()
    Dim myFSO As New FileSystemObject
    Dim myDS1 As Variant, myDS2 As Variant
    
    With myFSO.GetDrive("A")
        
        myDS1 = .TotalSize
        
        myDS2 = .AvailableSpace
    
    End With
    
    MsgBox "�w�ϥΪŶ��G" & Format(myDS1 - myDS2, "#,##0") & vbCrLf & _
        "�i�ΪŶ��G" & Format(myDS2, "#,##0") & vbCrLf & vbCrLf & _
        "�e�q�G" & Format(myDS1, "#,##0") & vbCrLf
End Sub


'--------------------------------------
'�d��85
'�d�ߺϺЪ�����
'--------------------------------------

Sub FSOSample4()
    Dim myFSO As New FileSystemObject
    Dim myDrv As Drive
    Dim myMsg As String
    
    For Each myDrv In myFSO.Drives
        
        myMsg = myMsg & myDrv.DriveLetter & "�G"
        
        Select Case myDrv.DriveType
            Case 0
                myMsg = myMsg & "����" & vbCrLf
            Case 1
                myMsg = myMsg & "������Ϻ�" & vbCrLf
            Case 2
                myMsg = myMsg & "�w���Ϻ�" & vbCrLf
            Case 3
                myMsg = myMsg & "�����Ϻ�" & vbCrLf
            Case 4
                myMsg = myMsg & "CD-ROM" & vbCrLf
            Case 5
                myMsg = myMsg & "RAM Disk" & vbCrLf
        End Select
    Next
    
    MsgBox myMsg
End Sub


'--------------------------------------
'�d��86
'�ˬd�Ϻо����ǳƪ��A
'--------------------------------------

Sub FSOSample5()
    Dim myFSO As New FileSystemObject
    
    If myFSO.Drives("A").IsReady = True Then
        FileCopy "C:\Excel2003VBA���νg\Fuji.txt", "A:\Fuji.txt"
    Else
        MsgBox "�����J�ϺФ�"
    End If
End Sub


'--------------------------------------
'�d��87
'���o�l��Ƨ�
'--------------------------------------

Sub FSOSample6()
    Dim myFSO As New FileSystemObject
    Dim myFld As Folder
    Dim i As Integer
    
    Worksheets("Sheet2").Activate
    i = 1
    
    With myFSO.GetFolder("C:\WINNT")
    
    '�Y�o�Ϳ��~�A�бN�W�泯�z���[�W���O�A�M�����U�泯�z��
    'With myFSO.GetFolder("C:\Windows")
    
        For Each myFld In .SubFolders
        
            i = i + 1
            Cells(i, 1).Value = myFld.Name
            
        Next
    End With
End Sub


'------------------------------------------------------
'�d��88
'���o��Ƨ��Ψ�l��Ƨ������ɮ��`�e�q
'------------------------------------------------------

Sub FSOSample7()
    Dim myFSO As New FileSystemObject
    Dim mySize1 As Variant, mySize2 As Variant
    
    With myFSO.GetFolder("C:\My Documents")
    
        mySize1 = .Size
        mySize2 = mySize1 / 1024 / 1024
        
        MsgBox _
            "C:\My Documents��Ƨ��������ɮפj�p���G" & vbCrLf & vbCrLf & _
            Format(mySize2, "#,##0.0") & "MB" & " (" & _
            Format(mySize1, "#,##0") & "Byte)"
    
    End With
End Sub




