Attribute VB_Name = "Module2"
Option Explicit


'------------------------------------------
'範例67
'資料數列(1)
'------------------------------------------

Sub SeriesSample()
    ActiveSheet.ChartObjects(1).Activate
    
    With ActiveChart.SeriesCollection(3)   '---參照對象為第3個資料數列
        .ChartType = xlLineMarkers          '---變更圖表類型
        .AxisGroup = xlSecondary            '---變更為副座標軸
    End With
End Sub


'------------------------------------------
'範例68
'資料數列(2)
'------------------------------------------

Sub SeriesCollectionSample()
    Dim mySeries As Series
    Dim myFormula As Variant
    Dim myMsg As String
    
    ActiveSheet.ChartObjects(1).Activate
   
    For Each mySeries In ActiveChart.SeriesCollection   'mySeries為SeriesCollection的成員
        myFormula = Split(mySeries.Formula, ",")        '以Split函數將參照範圍置入陣列內
        '將變數myMsg放入數列名稱及資料的參照範圍內
        myMsg = myMsg & mySeries.Name & " : " & myFormula(2) & vbCrLf
    Next
    
    MsgBox myMsg
End Sub


'------------------------------------------
'範例69
'圖表類型群組
'------------------------------------------

Sub ChartGroupsSample()
    ActiveSheet.ChartObjects(1).Activate
   
    With ActiveChart.ChartGroups(1)     '對第一個圖表類型群組進行下列變更
        .Overlap = 100                  '設定直條的重疊比例為100%
        .GapWidth = 50                  '將直條的間距設為欄寬的50%
    End With
End Sub


'------------------------------------------
'範例70
'資料點
'------------------------------------------

Sub PointsSample()
    Dim myFormula As Variant, myRange As Range, i As Long
    
    ActiveSheet.ChartObjects(1).Activate
   
    With ActiveChart.SeriesCollection(2)    '參照對象為第2個資料數列
        .MarkerSize = 5                     '將資料數列的標記大小設為5點
        myFormula = Split(.Formula, ",")    '將參照資料的範圍置入陣列
        
        Set myRange = Range(myFormula(2))   '設定資料將參照的儲存格範圍
        
        For i = 1 To myRange.Count
            '在參照資料中找到最大值時脫離迴圈
            'myRange.Cells(i)是.Points(i)的對應項
            If myRange.Cells(i).Value = _
                Application.WorksheetFunction.Max(myRange) Then
                Exit For
            End If
        Next
        
        .Points(i).MarkerSize = 10          '變更最大值資料點的標記大小
    End With
End Sub
