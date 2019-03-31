VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AddChart()
    Dim myChart As Chart
    Set myChart = Charts.Add
    With myChart
        .SetSourceData Source:=Sheet1.Range("A1:E3"), PlotBy:=xlRows
        .ChartType = xlColumnClustered 'TNND,谁能看出来这里面的l是1还是l！！！！
        .HasDataTable = True
        .ApplyDataLabels
        .HasTitle = True
        .ChartTitle.Text = "产品销量"
        .HasLegend = False
        .Name = "产品销量表"
    End With
    Set myChart = Nothing
End Sub

