VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AddChartObject()
    Dim objChart As ChartObject
    Set objChart = ActiveSheet.ChartObjects.Add(100, 100, 400, 300)
    With objChart.Chart
        .SetSourceData Source:=Sheet1.Range("A1:E3"), PlotBy:=xlColumns
        .ChartType = xlColumnClustered 'TNND,˭�ܿ������������lll11��11����l��������
        .ApplyDataLabels
        .HasTitle = True
        .ChartTitle.Text = "��Ʒ����"
    End With
    objChart.Name = "��Ʒ����ͼ��"
    Set objChart = Nothing
End Sub