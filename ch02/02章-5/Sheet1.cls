VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------
'範例18
'變更儲存格內容時執行事件程序
'---------------------------------------------------

Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    Dim r As Integer, myRange As Range
    
    Set myRange = Worksheets("客戶").Range("客戶編號")
    
    With Target
        '若變更儲存格D5的內容

        If .Row = 5 And .Column = 4 Then
    
            '取得客戶編號的位置
            r = Application.WorksheetFunction _
                .Match(Target.Value, myRange, 0)
    
            '在儲存格中顯示客戶名稱
            Range("F5") = Worksheets("客戶").Range("B1").Offset(r - 1).Value
        End If
    End With
End Sub

'---------------------------------------------------
'範例19
'變更選取範圍時執行事件程序
'---------------------------------------------------

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    With Target
        If .Row = 8 And .Column = 4 Then    '當選取儲存格D8時
            Range("E8").Select
        ElseIf .Row = 9 And .Column = 4 Then '當選取儲存格D9時
            Range("E9").Select
        ElseIf .Row = 10 And .Column = 4 Then '當選取儲存格D10時
            Range("E10").Select
        ElseIf .Row = 11 And .Column = 4 Then '當選取儲存格D11時
            Range("E11").Select
        End If
    End With
End Sub


