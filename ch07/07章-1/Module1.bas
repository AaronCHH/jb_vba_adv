Attribute VB_Name = "Module1"
Option Explicit

'------------------------------------------
'範例58
'以指定的資料範圍及數列建立圖表
'------------------------------------------

Sub MakeChart()
    Dim mySouce As Range
   
    Set mySouce = Range("B2").CurrentRegion
   
    '增加新的圖表
    Charts.Add
    
    '將數列的資料範圍指定為行
    ActiveChart.SetSourceData Source:=mySouce, PlotBy:=xlColumns
End Sub


'------------------------------------------
'範例59
'設定圖表標題文字
'------------------------------------------

Sub MakeChartTitle()
    Charts("Chart1").Activate
    
    With ActiveChart
        .HasTitle = True                    '顯示圖表標題
        .ChartTitle.Text = "九月份銷售量"      '設定標題文字
    End With
End Sub


'------------------------------------------
'範例60
'設定圖表標題位置
'------------------------------------------

Sub MoveChartTitle()
    Charts("Chart1").Activate
    
    With ActiveChart.ChartTitle
        .Top = 0                            '將標題上端位置設為圖表區域的上方
        .Left = 0                           '將標題左端位置設為圖表區域的左方
    End With
End Sub


'------------------------------------------
'範例61
'反轉橫條圖類別軸項目的排列順序
'------------------------------------------

Sub ReverseAxes()
    Worksheets("Sheet1").ChartObjects(1).Activate
    ActiveChart.Axes(xlCategory).ReversePlotOrder = True
End Sub


'------------------------------------------
'範例62
'設定座標軸標題
'------------------------------------------

Sub MakeAxisTitle()
    Worksheets("Sheet2").ChartObjects(1).Activate
    
    With ActiveChart
        .HasTitle = True                            '顯示圖表標題
        .ChartTitle.Text = "九月份銷售量"            '設定標題文字
        
        With .Axes(xlCategory)                      '類別X軸
            .HasTitle = True                        '顯示座標軸標題
            .AxisTitle.Text = "Office產品"    '設定座標軸標題文字
        End With
        
        With .Axes(xlValue)                         '數值Y軸
            .HasTitle = True                        '顯示座標軸標題
            .AxisTitle.Text = "數量"                '設定座標軸標題文字
        End With
    End With
End Sub


'------------------------------------------
'範例63
'切換格線的顯示/隱藏狀態
'------------------------------------------

Sub ToggleMajorGridlines()
    With ActiveSheet.ChartObjects(1).Chart.Axes(xlCategory)
        .HasMajorGridlines = Not .HasMajorGridlines
    End With
End Sub


'------------------------------------------
'範例64
'設定圖例
'------------------------------------------

Sub MakeLegend()
    With ActiveSheet.ChartObjects(1).Chart
        .ChartArea.AutoScaleFont = False            '固定字體大小
        
        .HasLegend = True                           '顯示圖例
        
        .Legend.Position = xlLegendPositionTop      '將圖例置於圖表區域上面的部份
        
        .PlotArea.Top = 0                           '將圖例置於繪圖區上端
        .PlotArea.Left = 0                          '將圖例置於繪圖區左端
        .PlotArea.Width = .ChartArea.Width          '設定繪圖區的寬度
        .PlotArea.Height = .ChartArea.Height        '設定繪圖區的高度
    End With
End Sub


'------------------------------------------
'範例65
'設定資料表格
'------------------------------------------

Sub MakeDataTable()
    With ActiveSheet.ChartObjects(1).Chart
        .HasLegend = False                          '由於圖例要附加在資料表格上，所以先不顯示圖例
        
        .HasDataTable = True                        '建立資料表格
        
        With .DataTable                             '資料表格的規格
            .ShowLegendKey = True                   '將圖例加到資料表格上
            .Font.Bold = True                       '將字體設為粗體
        End With
    End With
End Sub


'------------------------------------------
'範例66
'在XY散佈圖中顯示資料標籤
'------------------------------------------

Sub MakeDataLabels()
    Dim myRange As Range
    Dim i As Long

    '指定作為資料標籤的儲存格範圍(產品名稱)
    Set myRange = Range("A2", Range("A2").End(xlDown))
   
    ActiveSheet.ChartObjects(1).Activate
    
    '設定資料標籤
    ActiveChart.ApplyDataLabels
    
   '逐一變更DataLabel物件的Text屬性
    For i = 1 To myRange.Count
        ActiveChart.SeriesCollection(1).Points(i). _
            DataLabel.Text = myRange.Cells(i).Value
    Next i
End Sub
