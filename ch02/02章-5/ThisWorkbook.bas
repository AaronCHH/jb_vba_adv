Option Explicit

'---------------------------------------------------
'½d¨Ò18
'ÅÜ§óÀx¦s®æ¤º®e®É°õ¦æ¨Æ¥óµ{§Ç
'---------------------------------------------------

Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    Dim r As Integer, myRange As Range
    
    Set myRange = Worksheets("«È¤á").Range("«È¤á½s¸¹")
    
    With Target
        '­YÅÜ§óÀx¦s®æD5ªº¤º®e

        If .Row = 5 And .Column = 4 Then
    
            '¨ú±o«È¤á½s¸¹ªº¦ì¸m
            r = Application.WorksheetFunction _
                .Match(Target.Value, myRange, 0)
    
            '¦bÀx¦s®æ¤¤Åã¥Ü«È¤á¦WºÙ
            Range("F5") = Worksheets("«È¤á").Range("B1").Offset(r - 1).Value
        End If
    End With
End Sub

'---------------------------------------------------
'½d¨Ò19
'ÅÜ§ó¿ï¨ú½d³ò®É°õ¦æ¨Æ¥óµ{§Ç
'---------------------------------------------------

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    With Target
        If .Row = 8 And .Column = 4 Then    '·í¿ï¨úÀx¦s®æD8®É
            Range("E8").Select
        ElseIf .Row = 9 And .Column = 4 Then '·í¿ï¨úÀx¦s®æD9®É
            Range("E9").Select
        ElseIf .Row = 10 And .Column = 4 Then '·í¿ï¨úÀx¦s®æD10®É
            Range("E10").Select
        ElseIf .Row = 11 And .Column = 4 Then '·í¿ï¨úÀx¦s®æD11®É
            Range("E11").Select
        End If
    End With
End Sub


