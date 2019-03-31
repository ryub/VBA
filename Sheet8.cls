VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Add_Line()
    Dim chtChart As Chart, shpLine As Shape
    Dim AxisX As Axis, AxisY As Axis
    Dim x1 As Double, x2 As Double, y As Double, b As Double
    Dim sngAv As Single
    Const sLineName As String = "Line_In_Chart"
    Set chtChart = ActiveSheet.ChartObjects(1).Chart
        sngAv = Application.WorksheetFunction.Average(SeriesRange(1, chtChart))
    Set AxisX = chtChart.Axes(xlCategory, xlPrimary)
    Set AxisY = chtChart.Axes(xlValue, xlPrimary)
    With AxisY
        x1 = .Left + .Width
        x2 = x1 + AxisX.Width
            b = (.MaximumScale - sngAv) * .Height / (.MaximumScale - .MinimumScale)
        y = b + chtChart.PlotArea.Top
    End With
    On Error Resume Next
    chtChart.Shapes(sLineName).Delete
    Set shpLine = chtChart.Shapes.AddLine(x1, y, x2, y)
    With shpLine
        .Name = sLineName:
        .Line.ForeColor.ObjectThemeColor = 6
        .Line.Weight = 2
    End With
    Set shpLine = Nothing
    Set AxisX = Nothing
    Set AxisY = Nothing
    Set chtChart = Nothing
End Sub

