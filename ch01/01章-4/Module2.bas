Attribute VB_Name = "Module2"
Option Explicit

'-----------------------------
'絛ㄒ9
'虫ま计ㄏノ璹ㄧ计
'-----------------------------

Function MyMsg(Result As Variant) As String
    Select Case Result
        Case Is > 80
            MyMsg = "纔"
        Case Is > 60
            MyMsg = "▆"
        Case Is > 40
            MyMsg = "ぃの"
        Case Else
            MyMsg = "惠"
    End Select
End Function

'-----------------------------
'絛ㄒ10
'狡计ま计ㄏノ璹ㄧ计
'-----------------------------

Function MyFindS(r As Range, s As String) As Long
    Dim i As Long           '才穓碝挡狦纗计
    Dim myCell As Range

    For Each myCell In r
        If InStr(myCell.Value, s) > 0 Then
            i = i + 1
        End If
    Next
    
    MyFindS = i
End Function

'----------------------------------------------
'絛ㄒ11
'羆﹚纗絛瞅计ㄏノ璹ㄧ计
'----------------------------------------------

Function MySum(r As Range) As Double
    Dim myCell As Range

    For Each myCell In r
        MySum = Val(myCell.Value) + MySum
    Next
End Function

'----------------------------------------------
'絛ㄒ12
'笆滦笲衡ㄏノ璹ㄧ计
'----------------------------------------------

Function MyKaihi(n As Integer) As Long
    Application.Volatile
    MyKaihi = Range("A2").Value * n
End Function

