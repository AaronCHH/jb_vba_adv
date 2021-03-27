Attribute VB_Name = "Module1"
Option Explicit

'------------------------------------------
'�d��58
'�H���w����ƽd��μƦC�إ߹Ϫ�
'------------------------------------------

Sub MakeChart()
    Dim mySouce As Range
   
    Set mySouce = Range("B2").CurrentRegion
   
    '�W�[�s���Ϫ�
    Charts.Add
    
    '�N�ƦC����ƽd����w����
    ActiveChart.SetSourceData Source:=mySouce, PlotBy:=xlColumns
End Sub


'------------------------------------------
'�d��59
'�]�w�Ϫ���D��r
'------------------------------------------

Sub MakeChartTitle()
    Charts("Chart1").Activate
    
    With ActiveChart
        .HasTitle = True                    '��ܹϪ���D
        .ChartTitle.Text = "�E����P��q"      '�]�w���D��r
    End With
End Sub


'------------------------------------------
'�d��60
'�]�w�Ϫ���D��m
'------------------------------------------

Sub MoveChartTitle()
    Charts("Chart1").Activate
    
    With ActiveChart.ChartTitle
        .Top = 0                            '�N���D�W�ݦ�m�]���Ϫ�ϰ쪺�W��
        .Left = 0                           '�N���D���ݦ�m�]���Ϫ�ϰ쪺����
    End With
End Sub


'------------------------------------------
'�d��61
'�����������O�b���ت��ƦC����
'------------------------------------------

Sub ReverseAxes()
    Worksheets("Sheet1").ChartObjects(1).Activate
    ActiveChart.Axes(xlCategory).ReversePlotOrder = True
End Sub


'------------------------------------------
'�d��62
'�]�w�y�жb���D
'------------------------------------------

Sub MakeAxisTitle()
    Worksheets("Sheet2").ChartObjects(1).Activate
    
    With ActiveChart
        .HasTitle = True                            '��ܹϪ���D
        .ChartTitle.Text = "�E����P��q"            '�]�w���D��r
        
        With .Axes(xlCategory)                      '���OX�b
            .HasTitle = True                        '��ܮy�жb���D
            .AxisTitle.Text = "Office���~"    '�]�w�y�жb���D��r
        End With
        
        With .Axes(xlValue)                         '�ƭ�Y�b
            .HasTitle = True                        '��ܮy�жb���D
            .AxisTitle.Text = "�ƶq"                '�]�w�y�жb���D��r
        End With
    End With
End Sub


'------------------------------------------
'�d��63
'������u�����/���ê��A
'------------------------------------------

Sub ToggleMajorGridlines()
    With ActiveSheet.ChartObjects(1).Chart.Axes(xlCategory)
        .HasMajorGridlines = Not .HasMajorGridlines
    End With
End Sub


'------------------------------------------
'�d��64
'�]�w�Ϩ�
'------------------------------------------

Sub MakeLegend()
    With ActiveSheet.ChartObjects(1).Chart
        .ChartArea.AutoScaleFont = False            '�T�w�r��j�p
        
        .HasLegend = True                           '��ܹϨ�
        
        .Legend.Position = xlLegendPositionTop      '�N�ϨҸm��Ϫ�ϰ�W��������
        
        .PlotArea.Top = 0                           '�N�ϨҸm��ø�ϰϤW��
        .PlotArea.Left = 0                          '�N�ϨҸm��ø�ϰϥ���
        .PlotArea.Width = .ChartArea.Width          '�]�wø�ϰϪ��e��
        .PlotArea.Height = .ChartArea.Height        '�]�wø�ϰϪ�����
    End With
End Sub


'------------------------------------------
'�d��65
'�]�w��ƪ��
'------------------------------------------

Sub MakeDataTable()
    With ActiveSheet.ChartObjects(1).Chart
        .HasLegend = False                          '�ѩ�Ϩҭn���[�b��ƪ��W�A�ҥH������ܹϨ�
        
        .HasDataTable = True                        '�إ߸�ƪ��
        
        With .DataTable                             '��ƪ�檺�W��
            .ShowLegendKey = True                   '�N�Ϩҥ[���ƪ��W
            .Font.Bold = True                       '�N�r��]������
        End With
    End With
End Sub


'------------------------------------------
'�d��66
'�bXY���G�Ϥ���ܸ�Ƽ���
'------------------------------------------

Sub MakeDataLabels()
    Dim myRange As Range
    Dim i As Long

    '���w�@����Ƽ��Ҫ��x�s��d��(���~�W��)
    Set myRange = Range("A2", Range("A2").End(xlDown))
   
    ActiveSheet.ChartObjects(1).Activate
    
    '�]�w��Ƽ���
    ActiveChart.ApplyDataLabels
    
   '�v�@�ܧ�DataLabel����Text�ݩ�
    For i = 1 To myRange.Count
        ActiveChart.SeriesCollection(1).Points(i). _
            DataLabel.Text = myRange.Cells(i).Value
    Next i
End Sub
