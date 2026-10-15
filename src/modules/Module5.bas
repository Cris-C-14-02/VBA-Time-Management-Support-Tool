Attribute VB_Name = "Module5"
Option Explicit

Public Sub ClearSchoolTasks()
    DeleteCategoryTasks "School Work"
End Sub
Public Sub ClearHealthTasks()
    DeleteCategoryTasks "Health"
End Sub
Public Sub ClearSocialTasks()
    DeleteCategoryTasks "Social"
End Sub
Public Sub ClearDailyTasks()
    DeleteCategoryTasks "Daily"
End Sub
Public Sub ClearWorkFinanceTasks()
    DeleteCategoryTasks "Work-Finance"
End Sub
Public Sub ClearExtracurricularsTasks()
    DeleteCategoryTasks "Extracurriculars"
End Sub
Public Sub ClearUpcomingAndToDoTasks()
    ClearRows AppSheet("To-do and Sleep"), 3, 1, 4
End Sub
Public Sub ClearMasterList()
    ClearRows AppSheet("Master List"), 2, 1, 6
    TaskViews.RefreshAllTaskViews
End Sub
