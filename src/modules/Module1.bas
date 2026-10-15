Attribute VB_Name = "Module1"
Option Explicit

Public Sub SchoolWorkSheet()
    AppSheet("School Work").Activate
End Sub
Public Sub HealthSheet()
    AppSheet("Health").Activate
End Sub
Public Sub SocialSheet()
    AppSheet("Social").Activate
End Sub
Public Sub DailySheet()
    AppSheet("Daily").Activate
End Sub
Public Sub WorkFinanceSheet()
    AppSheet("Work-Finance").Activate
End Sub
Public Sub ExtracurricularsSheet()
    AppSheet("Extracurriculars").Activate
End Sub
Public Sub MasterListSheet()
    AppSheet("Master List").Activate
End Sub
Public Sub GraphsSheet()
    AppSheet("To-do and Sleep").Activate
End Sub
Public Sub TodoSheet()
    AppSheet("To-do and Sleep").Activate
End Sub
Public Sub MentalHealth()
    MentalHealthForm.Show
End Sub
Public Sub AddTask()
    AddTaskForm.Show
End Sub
Public Sub RefreshDashboard()
    On Error GoTo Failed
    TaskViews.RefreshAllTaskViews
    SleepService.UpdateSleepSummary
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub
