Attribute VB_Name = "Ä£¿é1"
Function SeriesRange(SeriesIndex As Integer, cht As Chart) As Range
    Dim chtSeries As Series
    Set chtSeries = cht.SeriesCollection(SeriesIndex)
    Set SeriesRange = Range(VBA.Split(VBA.Split(chtSeries.Formula, ",")(2), "!")(1))
    Set chtSeries = Nothing
End Function

