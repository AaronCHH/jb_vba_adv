Attribute VB_Name = "Module4"
Option Explicit
Option Base 1

'-------------------------
'�d��30
'�G���}�C���򥻨ϥΤ覡
'-------------------------

Sub SampleMatrix()
    Dim myData(3, 2) As String
    
    Dim i As Integer, j As Integer

    myData(1, 1) = "���ʼw"
    myData(1, 2) = "���֨T��"

    myData(2, 1) = "��ʹ�"
    myData(2, 2) = "���B��T"

    myData(3, 1) = "�^�ͳ�"
    myData(3, 2) = "�j�I�u�{"

    For i = 1 To 3
        For j = 1 To 2
            Debug.Print myData(i, j)
        Next j
    Next i
End Sub


'--------------------------------------------------
'�d��31
'�b�G���}�C���ϥ�For Each...Next���z��
'--------------------------------------------------

Sub SampleMatrix2()
    Dim myData(3, 2) As String
    
    Dim myVal As Variant

    myData(1, 1) = "���ʼw"
    myData(1, 2) = "���֨T��"

    myData(2, 1) = "��ʹ�"
    myData(2, 2) = "���B��T"

    myData(3, 1) = "�^�ͳ�"
    myData(3, 2) = "�j�I�u�{"

    For Each myVal In myData
        Debug.Print myVal
    Next
End Sub


'--------------------------------------------------
'�d��32
'�NVariant���A�ܼƳ]���x�s��d�򪺭�
'--------------------------------------------------

Sub RangeToVariant()
    Dim myData As Variant
    
    Dim r As Integer, c As Integer
    
    Worksheets("�ƥ��ӷ�").Activate
    
    myData = Range("A1").CurrentRegion.Value

    
    r = UBound(myData, 1)
    c = UBound(myData, 2)
    
    Worksheets("�ƥ�").Activate
    
    Range(Cells(1, 1), Cells(r, c)).Value = myData
End Sub
