Attribute VB_Name = "Module2"
Option Explicit


'------------------------------------------
'�d��67
'��ƼƦC(1)
'------------------------------------------

Sub SeriesSample()
    ActiveSheet.ChartObjects(1).Activate
    
    With ActiveChart.SeriesCollection(3)   '---�ѷӹ�H����3�Ӹ�ƼƦC
        .ChartType = xlLineMarkers          '---�ܧ�Ϫ�����
        .AxisGroup = xlSecondary            '---�ܧ󬰰Ʈy�жb
    End With
End Sub


'------------------------------------------
'�d��68
'��ƼƦC(2)
'------------------------------------------

Sub SeriesCollectionSample()
    Dim mySeries As Series
    Dim myFormula As Variant
    Dim myMsg As String
    
    ActiveSheet.ChartObjects(1).Activate
   
    For Each mySeries In ActiveChart.SeriesCollection   'mySeries��SeriesCollection������
        myFormula = Split(mySeries.Formula, ",")        '�HSplit��ƱN�ѷӽd��m�J�}�C��
        '�N�ܼ�myMsg��J�ƦC�W�٤θ�ƪ��ѷӽd��
        myMsg = myMsg & mySeries.Name & " : " & myFormula(2) & vbCrLf
    Next
    
    MsgBox myMsg
End Sub


'------------------------------------------
'�d��69
'�Ϫ������s��
'------------------------------------------

Sub ChartGroupsSample()
    ActiveSheet.ChartObjects(1).Activate
   
    With ActiveChart.ChartGroups(1)     '��Ĥ@�ӹϪ������s�նi��U�C�ܧ�
        .Overlap = 100                  '�]�w���������|��Ҭ�100%
        .GapWidth = 50                  '�N���������Z�]����e��50%
    End With
End Sub


'------------------------------------------
'�d��70
'����I
'------------------------------------------

Sub PointsSample()
    Dim myFormula As Variant, myRange As Range, i As Long
    
    ActiveSheet.ChartObjects(1).Activate
   
    With ActiveChart.SeriesCollection(2)    '�ѷӹ�H����2�Ӹ�ƼƦC
        .MarkerSize = 5                     '�N��ƼƦC���аO�j�p�]��5�I
        myFormula = Split(.Formula, ",")    '�N�ѷӸ�ƪ��d��m�J�}�C
        
        Set myRange = Range(myFormula(2))   '�]�w��ƱN�ѷӪ��x�s��d��
        
        For i = 1 To myRange.Count
            '�b�ѷӸ�Ƥ����̤j�Ȯɲ����j��
            'myRange.Cells(i)�O.Points(i)��������
            If myRange.Cells(i).Value = _
                Application.WorksheetFunction.Max(myRange) Then
                Exit For
            End If
        Next
        
        .Points(i).MarkerSize = 10          '�ܧ�̤j�ȸ���I���аO�j�p
    End With
End Sub
