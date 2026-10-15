Attribute VB_Name = "TaskCharts"
Option Explicit

Public Sub UpdateTaskChart(ByVal category As String)
    Dim ws As Worksheet, co As ChartObject, lastRow As Long, dateCol As Long
    Set ws = AppSheet(category)
    If category = "School Work" Then dateCol = 4 Else dateCol = 3
    lastRow = LastDataRow(ws, 1, dateCol + 1)
    On Error Resume Next
    Set co = ws.ChartObjects("TaskDuration")
    On Error GoTo 0
    If lastRow < 2 Then
        If Not co Is Nothing Then co.Delete
        Exit Sub
    End If
    If co Is Nothing Then Set co = ws.ChartObjects.Add(ws.Cells(1, dateCol + 3).Left, 24, 480, 280)
    co.Name = "TaskDuration"
    Do While co.Chart.SeriesCollection.Count > 0: co.Chart.SeriesCollection(1).Delete: Loop
    With co.Chart
        .ChartType = xlLineMarkers
        With .SeriesCollection.NewSeries
            .Name = "Estimated hours"
            .XValues = ws.Range(ws.Cells(2, dateCol), ws.Cells(lastRow, dateCol))
            .Values = ws.Range(ws.Cells(2, dateCol + 1), ws.Cells(lastRow, dateCol + 1))
            .Format.Line.ForeColor.RGB = RGB(13, 148, 136)
        End With
        .HasTitle = True: .ChartTitle.Text = category & ": estimated task hours"
        .HasLegend = False
        .Axes(xlCategory).CategoryType = xlCategoryScale
        .Axes(xlCategory).ReversePlotOrder = True
        .Axes(xlCategory).TickLabels.NumberFormat = "mmm d"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Estimated hours per task"
    End With
End Sub
