VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Test_SeriesRange()
    Dim chtChart As Chart
    Set chtChart = ActiveSheet.ChartObjects(1).Chart
    MsgBox SeriesRange(2, chtChart).Address
    Set chtChart = Nothing
End Sub

