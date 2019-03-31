VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AddChartObject()
    Dim objChart As ChartObject
    Set objChart = ActiveSheet.ChartObjects.Add(100, 100, 400, 300)
    With objChart.Chart
        .SetSourceData Source:=Sheet1.Range("A1:E3"), PlotBy:=xlRows
        .ChartType = xlColumnClustered 'TNND,谁能看出来这里面的lll11是11还是l！！！！
        .ApplyDataLabels
        .HasTitle = True
        .ChartTitle.Text = "产品销量"
    End With
    objChart.Name = "产品销量图表"
    Set objChart = Nothing
End Sub
