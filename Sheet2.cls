VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AddNewSeries()
    Dim chtChart As Chart, chtSeries As Series
    Dim ent As Integer, i As Integer
    Dim sngAv As Single, sStr As String
    Set chtChart = ActiveSheet.ChartObjects(1).Chart
        sngAv = Application.WorksheetFunction.Average(SeriesRange(1, chtChart))
    cnt = chtChart.SeriesCollection(1).Points.Count
    For i = 1 To cnt
    sStr = sStr & "," & sngAv
    Next
    sStr = Mid(sStr, 2, Len(sStr))
    sStr = "={" & sStr & "}"
    Set chtSeries = chtChart.SeriesCollection.NewSeries
    chtSeries.ChartType = xlLine
    chtSeries.Values = sStr
    Set chtSeries = Nothing
    Set chtChart = Nothing
 End Sub
