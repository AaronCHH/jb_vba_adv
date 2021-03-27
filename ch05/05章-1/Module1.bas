Attribute VB_Name = "Module1"
Option Explicit

'-----------------------------------
'�d��36
'�ܧ�Excel���D�C����r
'-----------------------------------

Sub ChangeTitleBar()
    Application.Caption = "�X�ХX��(��)���q"
    ActiveWindow.Caption = "���ƪ��t��"
End Sub


'----------------------------------------
'�d��37
'�N�����K������
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
'�d��38
'�NExcel���ù���ܰϰ줧�~
'----------------------------------------

Sub HideExcel()
    Dim myTop As Double, myLeft As Double
    
    Application.WindowState = xlNormal
    
    myTop = Application.Top
    myLeft = Application.Left
    
    MsgBox "����Excel"
    Application.Left = -Application.Width
    
    MsgBox "�A�����Excel"
    Application.Top = myTop
    Application.Left = myLeft
End Sub


'----------------------------------------
'�d��39
'�����ù���s�\��
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
'�d��40
'���çR���u�@���T�{�T��
'----------------------------------------

Sub StopAlertMsg()
    Application.DisplayAlerts = False

    Worksheets.Add.Name = "Dummy"
    MsgBox "�s�W�u�@��uDummy�v" & vbCrLf & _
        "���ۡA�R���uDummy�v"

    Worksheets("Dummy").Delete
    
    Application.DisplayAlerts = True
End Sub


'----------------------------------------
'�d��41
'�b���A�C��ܰT��
'----------------------------------------

Sub DispStatusBar()
    Dim myStatusBar As Boolean
    Dim myCell As Range
    
    Worksheets("Sheet3").Activate
    
    myStatusBar = Application.DisplayStatusBar
    
    Application.DisplayStatusBar = True
    
    For Each myCell In Range("A1:C5")
        myCell.Value = "ABC"
        
        Application.StatusBar = "�g�J " & myCell.Address & " ��"

        Application.Wait Now + TimeValue("00:00:01")
    Next myCell
   
    Application.StatusBar = False
        
    Application.DisplayStatusBar = myStatusBar
End Sub


'----------------------------------------
'�d��42
'���ع�ܤ��������P�_
'----------------------------------------

Sub DispBuiltinDialog()
    Dim myRtn As Boolean
    
    myRtn = Application.Dialogs(xlDialogOptionsView).Show
    
    If myRtn = False Then
        MsgBox "�I��u�����v" & vbCrLf & _
            "�B�z����"
        Exit Sub
    End If

    MsgBox "�~��i��B�z"
'
' �i��B�z
'
End Sub


'----------------------------------------------
'�d��43
'�ϥ�GetOpenFilename��k�}�Ҥ�r�ɮ�
'----------------------------------------------

Sub OpenTxtFile()
    Dim myFName As String
    
    myFName = Application. _
        GetOpenFilename("��r�ɮ�(*.prn; *.txt; *.csv),*.prn;*.txt;*.csv")
    
    If myFName <> "False" Then
        Workbooks.OpenText Filename:=myFName, Comma:=True
    End If
End Sub

