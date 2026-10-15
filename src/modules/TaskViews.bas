Attribute VB_Name = "TaskViews"
Option Explicit

Public Sub RefreshAllTaskViews()
    Dim category As Variant
    ValidateTaskSource
    For Each category In TaskCategories()
        RefreshCategoryView CStr(category)
    Next category
    RefreshTodayTasks
End Sub

Public Sub RefreshCategoryView(ByVal category As String)
    Dim source As Worksheet, destination As Worksheet, row As Long, targetRow As Long, dateCol As Long
    If Not IsTaskCategory(category) Then Err.Raise 5, , "Invalid task category."
    ValidateTaskSource
    Set source = AppSheet("Master List"): Set destination = AppSheet(category)
    If category = "School Work" Then dateCol = 4 Else dateCol = 3
    ClearRows destination, 2, 1, IIf(category = "School Work", 6, 4)
    destination.Range(destination.Cells(2, 1), destination.Cells(Application.Max(2, LastDataRow(source, 1, 6)), IIf(category = "School Work", 3, 2))).NumberFormat = "@"
    targetRow = 2
    For row = 2 To LastDataRow(source, 1, 6)
        If CStr(source.Cells(row, 1).value) = category Then
            If category = "School Work" Then
                destination.Range("A" & targetRow & ":F" & targetRow).value = source.Range("A" & row & ":F" & row).value
            Else
                destination.Cells(targetRow, 1).value = category
                destination.Cells(targetRow, 2).value = source.Cells(row, 3).value
                destination.Cells(targetRow, 3).value = source.Cells(row, 4).value
                destination.Cells(targetRow, 4).value = source.Cells(row, 5).value
            End If
            targetRow = targetRow + 1
        End If
    Next row
    If targetRow > 2 Then
        destination.Range(destination.Cells(1, 1), destination.Cells(targetRow - 1, IIf(category = "School Work", 6, 4))).Sort Key1:=destination.Cells(2, dateCol), Order1:=xlDescending, Header:=xlYes
        destination.Range(destination.Cells(2, dateCol), destination.Cells(targetRow - 1, dateCol)).NumberFormat = "yyyy-mm-dd"
        destination.Range(destination.Cells(2, dateCol + 1), destination.Cells(targetRow - 1, dateCol + 1)).NumberFormat = "0.0"
        If category = "School Work" Then destination.Range("F2:F" & targetRow - 1).NumberFormat = "yyyy-mm-dd"
    End If
    TaskCharts.UpdateTaskChart category
End Sub

Public Sub RefreshTodayTasks()
    Dim source As Worksheet, output As Worksheet, row As Long, todayRow As Long, upcomingRow As Long, starts As Date
    ValidateTaskSource
    Set source = AppSheet("Master List"): Set output = AppSheet("To-do and Sleep")
    ClearRows output, 3, 1, 4
    output.Range("A3:A" & Application.Max(3, LastDataRow(source, 1, 6))).NumberFormat = "@"
    output.Range("C3:C" & Application.Max(3, LastDataRow(source, 1, 6))).NumberFormat = "@"
    output.Range("B2").value = Date: output.Range("B2").NumberFormat = "yyyy-mm-dd"
    output.Range("C1").value = "Upcoming (next 7 days)"
    todayRow = 3: upcomingRow = 3
    For row = 2 To LastDataRow(source, 1, 6)
        If IsDate(source.Cells(row, 4).value) Then
            starts = DateValue(source.Cells(row, 4).value)
            If starts = Date Then
                output.Cells(todayRow, 1).value = source.Cells(row, 3).value
                output.Cells(todayRow, 2).value = starts
                todayRow = todayRow + 1
            ElseIf starts > Date And starts <= Date + 7 Then
                output.Cells(upcomingRow, 3).value = source.Cells(row, 3).value
                output.Cells(upcomingRow, 4).value = starts
                upcomingRow = upcomingRow + 1
            End If
        End If
    Next row
    If todayRow > 3 Then output.Range("B3:B" & todayRow - 1).NumberFormat = "yyyy-mm-dd"
    If upcomingRow > 3 Then
        output.Range("D3:D" & upcomingRow - 1).NumberFormat = "yyyy-mm-dd"
        output.Range("C2:D" & upcomingRow - 1).Sort Key1:=output.Range("D3"), Order1:=xlAscending, Header:=xlYes
    End If
    With output.Range("A3:D" & Application.Max(3, todayRow - 1, upcomingRow - 1))
        .WrapText = True
        .Rows.AutoFit
    End With
End Sub
