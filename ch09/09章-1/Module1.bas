Attribute VB_Name = "Module1"
Option Explicit
    Dim objWord As New Word.Application     'Word���ε{��
    Dim objWordDoc As Word.Document         '�s�W��Word���

'-----------------------------------------------------------
'�d��93
'�ϥΦ���ô������Word����
'-----------------------------------------------------------

Sub CreateWordApp()
    Dim myPath As String
    Dim myReturn As Integer
    
    myPath = ActiveWorkbook.Path & "\"
    
    myReturn = MsgBox(prompt:="�uTest.doc�v�O�������ܡH" & _
        "(�Y�O�}�Ҫ��A�N�o�Ϳ��~)", Buttons:=vbYesNo)
    If myReturn = vbNo Then Exit Sub
    
    With objWord
        .Visible = True                                               '���Word
        .WindowState = wdWindowStateMaximize '�N�����̤j��
        .Documents.Add                                            '�s�W���
        
        '�N�s�W���N�J�����ܼƤ�
        Set objWordDoc = .ActiveDocument
    End With
    
    '�b��󤤴��J��r
    With objWord.Selection
        .InsertAfter "�إ�Word���󪺴���"
        .InsertParagraphAfter
        .InsertAfter Now() & " �إ�"
        .MoveRight
    End With
    
    '�]�w�q��1���榡
    With objWordDoc.Paragraphs(1).Range
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Font
            .Name = "�з���"
            .Size = 20
            .Bold = True
        End With
    End With
    
    '�]�w�q��2���榡
    objWordDoc.Paragraphs(2).Range.ParagraphFormat _
        .Alignment = wdAlignParagraphRight
    
    Application.Wait Now() + TimeValue("00:00:03")
    
    objWord.WindowState = wdWindowStateMinimize '�N�����̤p��

    MsgBox "�Ұ�Word�ëإ߷s���" & Chr(13) & _
        "���U[�T�w]���s���x�s��������Word"

    objWordDoc.SaveAs myPath & "Test.doc"   '�N�s�W�����t�s�s��
    objWordDoc.Close                        '�����s�W�����
    objWord.Quit                            '����Word

    Set objWord = Nothing                   '�M�������ܼƪ����e
    Set objWordDoc = Nothing
End Sub


'-----------------------------------------------------------
'�d��94
'�NExcel�u�@��ǰe��Word���
'-----------------------------------------------------------

Sub CreateWordApp2()
    
    With objWord
        .Visible = True                      '���Word
        .WindowState = wdWindowStateMaximize '�N�����̤j��
        .Documents.Open ActiveWorkbook.Path & "\Report.doc" '�}��Report.doc
        
        '�NReport.doc�N�J�����ܼƤ�
        Set objWordDoc = .ActiveDocument
        
        '�b��󤤴��J��r
        With .Selection
            .Move Count:=objWordDoc.Characters.Count
            .InsertParagraphAfter
            .InsertAfter "CD�c��i��"
            .InsertParagraphAfter
            .MoveRight
        End With
    End With
    
    '�ƻs�x�s�檺���
    Worksheets("CD�c��q").Range("CD�c��i��").Copy
    
    '�K��Word��
    With objWord.Selection
        .Paste
        .TypeParagraph
    End With
    
    '�ƻs�Ϫ�
    Worksheets("CD�c��q").ChartObjects(1).Copy
    
    '�]�w�K��Word���榡
    With objWord
        .Selection.PasteSpecial Placement:=wdInLine, _
            DataType:=wdPasteMetafilePicture
        .Selection.ParagraphFormat.Alignment = _
            wdAlignParagraphCenter
    End With
    
    '�C�L(�C�L�ɱN���_����������)
    objWord.PrintOut Background:=False
    
    '���x�s���A����Word
    objWordDoc.Close SaveChanges:=False
    
    objWord.Quit                '����Word
    
    Set objWord = Nothing       '�M�������ܼƪ����e
    Set objWordDoc = Nothing
End Sub


'----------------------------------
'�d��95
'�s���w�Ұʪ�Word
'----------------------------------

Sub GetWordApp()
    On Error GoTo HandleErr
    Dim myAppOpen As Boolean
    
    Set objWord = GetObject(, "Word.Application")
    myAppOpen = True
    
MacroContinue:
    If myAppOpen = False Then
        Set objWord = CreateObject("Word.Application")
    End If
    
    With objWord
        .Visible = True
        .WindowState = wdWindowStateMinimize
        .Documents.Add
    End With
    
    Set objWord = Nothing
    
    Exit Sub

HandleErr:
    '�Y�|���NActiveX����إ߬�����
    If Err.Number = 429 Then
        myAppOpen = False
        Resume MacroContinue
    End If
End Sub


