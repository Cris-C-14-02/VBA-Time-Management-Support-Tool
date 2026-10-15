Attribute VB_Name = "Module4"
Option Explicit

Public Sub NewestSchoolTaskOnTop()
    TaskViews.RefreshCategoryView "School Work"
End Sub
Public Sub NewestHealthTaskToTop()
    TaskViews.RefreshCategoryView "Health"
End Sub
Public Sub NewestSocailTaskToTop()
    TaskViews.RefreshCategoryView "Social"
End Sub
Public Sub NewestDailyTaskToTop()
    TaskViews.RefreshCategoryView "Daily"
End Sub
Public Sub NewestWorkFinanceTaskToTop()
    TaskViews.RefreshCategoryView "Work-Finance"
End Sub
Public Sub NewestECTaskToTop()
    TaskViews.RefreshCategoryView "Extracurriculars"
End Sub
