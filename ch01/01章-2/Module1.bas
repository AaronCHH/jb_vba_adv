Attribute VB_Name = "Module1"
Option Explicit

'---------------------------------------------
'�d��3
'�ϥΤޥΩI�s��L����ï���{��
'---------------------------------------------

Sub SansyoSettei()
    Call Sample1
End Sub

'------------------------------------------------
'�d��4
'�ϥ�Run��k�I�s��L����ï���{��
'(�����w���|)
'------------------------------------------------

Sub AppRun()
    Application.Run "Subrtin2.xls!Sample2"
End Sub

'-------------------------------------------------
'�d��5
'�ϥΫ��w�����|�I�s��L����ï���{��
'-------------------------------------------------

Sub AppRun2()
    Dim myWBPath As String
    
    myWBPath = ActiveWorkbook.Path
    
    Application.Run "'" & myWBPath & "\Subrtin2.xls'!Sample2"
End Sub

