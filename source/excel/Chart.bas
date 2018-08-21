Attribute VB_Name = "excel_Chart"
Option Explicit

Public Sub Chart_ClearSeries(ByVal Chart As Chart)
    Dim seriesCollection As seriesCollection: Set seriesCollection = Chart.seriesCollection()
    
    While seriesCollection.count > 1
        seriesCollection.Item(1).Delete
    Wend
    
    Set seriesCollection = Nothing
End Sub
