Attribute VB_Name = "SleepService"
Option Explicit

Public Function IsSleepHours(ByVal value As Variant, ByRef hours As Double) As Boolean
    If Not TryNumber(value, hours, True) Then Exit Function
    IsSleepHours = (hours <= 24)
End Function

Public Sub SaveSleep(ByVal day As String, ByVal hours As Double)
    Dim days As Variant, index As Long
    If hours < 0 Or hours > 24 Then Err.Raise 5, , "Sleep hours must be between 0 and 24."
    days = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    For index = 0 To 6
        If days(index) = day Then
            With AppSheet("To-do and Sleep").Cells(index + 8, 6)
                .value = hours: .NumberFormat = "0.0"
            End With
            UpdateSleepSummary
            Exit Sub
        End If
    Next index
    Err.Raise 5, , "Select a listed day of the week."
End Sub

Public Sub UpdateSleepSummary()
    Dim ws As Worksheet, row As Long, total As Double, days As Long, hours As Double
    Set ws = AppSheet("To-do and Sleep")
    For row = 8 To 14
        If Len(CStr(ws.Cells(row, 6).value)) > 0 Then
            If Not IsSleepHours(ws.Cells(row, 6).value, hours) Then Err.Raise 5, , "Invalid sleep entry at row " & row
            total = total + hours: days = days + 1
        End If
    Next row
    If days = 0 Then
        ws.Range("H19").value = "No sleep entries yet."
    Else
        ws.Range("H19").value = "Average sleep: " & Format$(total / days, "0.0") & " hours across " & days & " recorded day(s)."
    End If
End Sub
