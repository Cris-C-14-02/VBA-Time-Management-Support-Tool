Attribute VB_Name = "Module2"
Option Explicit

Public Sub DeleteGraphs()
    Dim co As ChartObject
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
    For Each co In ActiveSheet.ChartObjects
        If co.Name = "TaskDuration" Then co.Delete
    Next co
End Sub
Public Sub GraphSchoolWork()
    TaskCharts.UpdateTaskChart "School Work"
End Sub
Public Sub GraphHealth()
    TaskCharts.UpdateTaskChart "Health"
End Sub
Public Sub GraphSocial()
    TaskCharts.UpdateTaskChart "Social"
End Sub
Public Sub GraphDaily()
    TaskCharts.UpdateTaskChart "Daily"
End Sub
Public Sub GraphWorkFinance()
    TaskCharts.UpdateTaskChart "Work-Finance"
End Sub
Public Sub GraphECs()
    TaskCharts.UpdateTaskChart "Extracurriculars"
End Sub
