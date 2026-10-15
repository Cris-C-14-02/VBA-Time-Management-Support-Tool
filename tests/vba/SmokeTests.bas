Attribute VB_Name = "SmokeTests"
Option Explicit
Private Sub Require(ByVal condition As Boolean, ByVal description As String)
    If Not condition Then Err.Raise vbObjectError + 902, "SmokeTests", description
End Sub

Public Function RunSmokeTests() As String
    On Error GoTo Failed
    Dim d As Date, hours As Double, outsider As Workbook, savedRows As Long, failure As Long
    Dim stage As String
    stage = "Load AddTaskForm"
    Load AddTaskForm
    Require AddTaskForm.txtSubcategory.ListCount = 6, "Task categories did not initialize"
    stage = "Check course visibility"
    Require Not AddTaskForm.txtCourse.Visible, "School controls visible without school category"
    AddTaskForm.txtSubcategory.Value = "School Work"
    Require AddTaskForm.txtCourse.Visible, "School fields did not become visible"
    Unload AddTaskForm
    stage = "Load AddSleepForm"
    Load AddSleepForm
    Require AddSleepForm.cbDayOfWeek.ListCount = 7, "Sleep days did not initialize"
    Unload AddSleepForm
    stage = "Service tests"
    Require TryDateParts("2024", "2", "29", d), "Leap day rejected"
    Require Not TryDateParts("2023", "2", "29", d), "Invalid date accepted"
    Require Not TryNumber(0, hours), "Zero estimated hours accepted"
    Require Not TryNumber(-2, hours), "Negative estimated hours accepted"
    Require Not TryNumber("bad", hours), "Text hours accepted"
    Require IsSleepHours(0, hours), "Zero sleep rejected"
    Require IsSleepHours(24, hours), "24 hours rejected"
    Require Not IsSleepHours(24.1, hours), "Sleep over 24 accepted"
    ClearMasterList
    SaveTask "School Work", "DEMO101", "Course task", Date, 2, Date + 1
    SaveTask "Health", "Ignored", "Walk", Date, 0.5, Empty
    SaveTask "Health", "", "Tomorrow", Date + 1, 1, Empty
    SaveTask "Daily", "", "Seven days", Date + 7, 1, Empty
    SaveTask "Daily", "", "Eight days", Date + 8, 1, Empty
    Set outsider = Application.Workbooks.Add
    TaskViews.RefreshAllTaskViews
    Require AppSheet("To-do and Sleep").Range("A3").Value = "Course task", "Today's list missing task"
    Require Application.CountA(AppSheet("To-do and Sleep").Range("A3:A20")) = 2, "Today's list count wrong"
    Require Application.CountA(AppSheet("To-do and Sleep").Range("C3:C20")) = 2, "Next-seven-days boundary wrong"
    Require AppSheet("Health").Range("B2").Value = "Tomorrow", "Category not sorted newest first"
    Require AppSheet("Health").ChartObjects.Count = 1, "Category chart missing"
    TaskViews.RefreshAllTaskViews
    Require AppSheet("Health").ChartObjects.Count = 1, "Repeated refresh duplicates chart"
    Require Application.CountA(outsider.Worksheets(1).UsedRange) = 0, "Other workbook was modified"
    outsider.Close False
    savedRows = LastDataRow(AppSheet("Master List"), 1, 6)
    On Error Resume Next
    SaveTask "School Work", "DEMO101", "Bad due date", Date, 1, Date - 1
    failure = Err.Number: Err.Clear
    On Error GoTo Failed
    Require failure <> 0, "Earlier due date accepted"
    Require LastDataRow(AppSheet("Master List"), 1, 6) = savedRows, "Invalid task was written"
    DeleteCategoryTasks "Health"
    Require Application.CountA(AppSheet("Health").Range("A2:D20")) = 0, "Category deletion left data"
    Require AppSheet("Health").ChartObjects.Count = 0, "Category deletion left chart"
    Require Application.CountA(AppSheet("To-do and Sleep").Range("A3:A20")) = 1, "Today's list has stale rows"
    AppSheet("To-do and Sleep").Range("F8:F14").ClearContents
    SaveSleep "Sunday", 0
    SaveSleep "Monday", 8
    Require InStr(AppSheet("To-do and Sleep").Range("H19").Value, "4.0 hours across 2") > 0, "Missing days treated as zero or zero excluded"
    failure = 0
    On Error Resume Next
    SaveSleep "Invalid day", 8
    failure = Err.Number: Err.Clear
    On Error GoTo Failed
    Require failure <> 0, "Invalid sleep day accepted"
    ClearMasterList
    Require Application.CountA(AppSheet("To-do and Sleep").Range("A3:D20")) = 0, "Clearing master left views"
    RunSmokeTests = "PASS: date and duration validation; school due date; master-to-category synchronization; today and next-seven-days boundaries; sorting; deletion; chart repeatability; other-workbook isolation; sleep 0/24 boundaries and recorded-day average."
    Exit Function
Failed:
    RunSmokeTests = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function

Public Sub BuildDemo()
    Dim days As Variant, values As Variant, index As Long
    ClearMasterList
    SaveTask "School Work", "STA-DEMO", "Review statistics notes", Date, 2, Date + 2
    SaveTask "Health", "", "Evening walk", Date, 0.5, Empty
    SaveTask "Daily", "", "Plan tomorrow", Date, 0.3, Empty
    SaveTask "Work-Finance", "", "Review internship roles", Date + 1, 1, Empty
    SaveTask "School Work", "ECO-DEMO", "Prepare economics assignment", Date + 3, 3, Date + 5
    SaveTask "Social", "", "Meet study group", Date + 4, 1.5, Empty
    SaveTask "Extracurriculars", "", "Club project planning", Date + 6, 2.5, Empty
    TaskViews.RefreshAllTaskViews
    AppSheet("To-do and Sleep").Range("F8:F14").ClearContents
    days = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    values = Array(7.5, 8, 6.5, 7.5, 8, 7, 8)
    For index = 0 To 6: SaveSleep CStr(days(index)), CDbl(values(index)): Next index
End Sub
