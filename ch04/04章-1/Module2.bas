Attribute VB_Name = "Module2"
Option Explicit
Option Base 1

'---------------------
'�d��27
'�ʺA�}�C�ܼƪ��򥻻y�k
'---------------------

Sub ReDimSample()
    Dim myName() As String
    
    Dim r As Long, i As Long

    Worksheets("3�~A�Z").Activate
    
    '��X��ƿ�J���̫�@�C
    r = Range("A65536").End(xlUp).Row

    ReDim myName(r)
    
    For i = 1 To r
        myName(i) = Cells(i, 1).Value
    Next
        
    For i = 1 To r
        Debug.Print myName(i)
    Next
End Sub


'----------------------------------
'�d��28
'�ϥ�LBound��Ƥ�UBound��ƫإ߰j��
'----------------------------------

Sub ReDimSample2()
    Dim myName() As String
    
    Dim r As Long, i As Long

    Worksheets("3�~A�Z").Activate
    
    r = Range("A65536").End(xlUp).Row

    ReDim myName(r)
    
    For i = LBound(myName) To UBound(myName)
        myName(i) = Cells(i, 1).Value
    Next
        
    For i = LBound(myName) To UBound(myName)
        Debug.Print myName(i)
    Next
End Sub
