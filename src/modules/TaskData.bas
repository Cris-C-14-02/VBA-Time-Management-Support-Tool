Attribute VB_Name = "TaskData"
Option Explicit

Public Function TaskCategories() As Variant
    TaskCategories = Array("School Work", "Health", "Social", "Daily", "Work-Finance", "Extracurriculars")
End Function

Public Function IsTaskCategory(ByVal category As String) As Boolean
    Dim value As Variant
    For Each value In TaskCategories()
        If category = value Then IsTaskCategory = True: Exit Function
    Next value
End Function

Public Sub SaveTask(ByVal category As String, ByVal course As String, ByVal task As String, ByVal starts As Date, ByVal hours As Double, ByVal due As Variant)
    Dim ws As Worksheet, row As Long
    If Not IsTaskCategory(category) Or Len(Trim$(task)) = 0 Or hours <= 0 Then Err.Raise 5, , "Invalid task."
    If category = "School Work" Then
        If Len(Trim$(course)) = 0 Or Not IsDate(due) Then Err.Raise 5, , "School tasks require a course and due date."
        If CDate(due) < starts Then Err.Raise 5, , "Due date must be on or after start date."
    End If
    Set ws = AppSheet("Master List")
    row = LastDataRow(ws, 1, 6) + 1
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 3)).NumberFormat = "@"
    ws.Cells(row, 1).Value2 = category
    If category = "School Work" Then ws.Cells(row, 2).Value2 = Trim$(course)
    ws.Cells(row, 3).Value2 = Trim$(task)
    ws.Cells(row, 4).value = starts: ws.Cells(row, 4).NumberFormat = "yyyy-mm-dd"
    ws.Cells(row, 5).value = hours: ws.Cells(row, 5).NumberFormat = "0.0"
    If category = "School Work" Then ws.Cells(row, 6).value = CDate(due)
    ws.Cells(row, 6).NumberFormat = "yyyy-mm-dd"
End Sub

Public Sub ValidateTaskSource()
    Dim ws As Worksheet, row As Long, hours As Double, category As String
    Set ws = AppSheet("Master List")
    For row = 2 To LastDataRow(ws, 1, 6)
        If Application.CountA(ws.Range(ws.Cells(row, 1), ws.Cells(row, 6))) > 0 Then
            category = CStr(ws.Cells(row, 1).value)
            If Not IsTaskCategory(category) Then Err.Raise 5, , "Invalid category at Master List row " & row
            If Len(Trim$(CStr(ws.Cells(row, 3).value))) = 0 Or Not IsDate(ws.Cells(row, 4).value) Then Err.Raise 5, , "Invalid task or date at Master List row " & row
            If Not TryNumber(ws.Cells(row, 5).value, hours) Then Err.Raise 5, , "Invalid estimated hours at Master List row " & row
            If category = "School Work" Then
                If Len(Trim$(CStr(ws.Cells(row, 2).value))) = 0 Or Not IsDate(ws.Cells(row, 6).value) Then Err.Raise 5, , "Missing course or due date at Master List row " & row
                If CDate(ws.Cells(row, 6).value) < CDate(ws.Cells(row, 4).value) Then Err.Raise 5, , "Due date precedes start date at Master List row " & row
            End If
        End If
    Next row
End Sub

Public Sub DeleteCategoryTasks(ByVal category As String)
    Dim ws As Worksheet, row As Long
    If Not IsTaskCategory(category) Then Err.Raise 5, , "Invalid task category."
    Set ws = AppSheet("Master List")
    For row = LastDataRow(ws, 1, 6) To 2 Step -1
        If CStr(ws.Cells(row, 1).value) = category Then ws.Range(ws.Cells(row, 1), ws.Cells(row, 6)).Delete Shift:=xlUp
    Next row
    TaskViews.RefreshAllTaskViews
End Sub
