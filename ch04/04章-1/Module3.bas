Attribute VB_Name = "Module3"
Option Explicit
Option Base 1

'------------------------------------
'�d��29
'�O�d�}�C���ȡA�ܧ󤸯��ƶq
'------------------------------------

Sub PreserveSample()
    Dim myName() As String
    
    Dim r As Long, r2 As Long
    Dim i As Long

    '�N�u3�~A�Z�v�u�@������Ƹm�J�}�C�ܼƤ�
    Worksheets("3�~A�Z").Activate
    
    r = Range("A65536").End(xlUp).Row

    ReDim myName(r)
    
    For i = 1 To r
        myName(i) = Cells(i, 1).Value
    Next
        
    '�N�u3�~B�Z�v�u�@������Ʒs�W��}�C�ܼƤ�
    Worksheets("3�~B�Z").Activate
    
    r2 = Range("A65536").End(xlUp).Row

    ReDim Preserve myName(r + r2)
    
    For i = r + 1 To r + r2
        myName(i) = Cells(i - r, 1).Value
    Next
    
    '�N�}�C�ܼƪ��ȿ�X��Y�ɹB�������
    For i = LBound(myName) To UBound(myName)
        Debug.Print myName(i)
    Next
End Sub

