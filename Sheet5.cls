VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CreateChartObject2()
    Dim chtObj As ChartObject, ActSht As Worksheet, Ser As Series
    Dim sngLeft As Single, sngTop As Single
    Set ActSht = ActiveSheet
    sngLeft = ActSht.Range("D7").Left
    sngTop = ActSht.Range("D7").Top
    Set chtObj = ActSht.ChartObjects.Add(sngLeft, sngTop, 550, 300)
    With chtObj.Chart
        .SetSourceData ActSht.Range("A1：B6"), xlColumns
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "各店销售情况"
        With .Axes(xlValue, xlPrimary)
            .MaximumScale = 1000
            .MinimumScale = 0
            .MajorUnit = 100
            .HasTitle = True
            .AxisTitle.Text = ActSht.Range("B1")
        End With
        Set Ser = .SeriesCollection.NewSeries
        With Ser
            .Values = ActSht.Range("C2：C6")
            .ChartType = xlLineMarkers
            .Name = ActSht.Range("C1")
            .AxisGroup = xlSecondary
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 5
        End With
        With .Axes(xlValue, xlSecondary)
            .MaximumScale = 1
            .MinimumScale = 0
            .HasTitle = True
            .AxisTitle.Text = ActSht.Range("C1")
        End With
        .SetElement (msoElementPrimaryValueAxisTitleVertical)
        .SetElement (msoElementSecondaryValueAxisTitleVertic)
        .SetElement (msoE1ementLegendBo11om)
    End With
    Set chtObj = Nothing
    Set ActSht = Nothing
End Sub
