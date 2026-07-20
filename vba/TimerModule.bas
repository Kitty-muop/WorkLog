Attribute VB_Name = "TimerModule"
Option Explicit

Private mStartTime As Double

Sub TimerButton_Click()
    Dim btn As Button
    Dim btnName As String
    Dim targetRow As Long
    Dim endTime As Double

    On Error Resume Next
    btnName = Application.Caller
    If btnName = "" Then Exit Sub
    On Error GoTo 0

    If ActiveSheet.Name <> "Time Entries" Then
        MsgBox "Please use this button in the Time Entries sheet.", vbExclamation, "Wrong sheet"
        Exit Sub
    End If

    Set btn = ActiveSheet.Buttons(btnName)

    If btn.Caption = "Start Timer" Then
        mStartTime = Now
        btn.Caption = "Stop Timer"
        btn.Font.Bold = True
    Else
        If mStartTime = 0 Then
            MsgBox "Timer not started. Please click Start Timer first.", vbExclamation, "Error"
            Exit Sub
        End If

        endTime = Now
        targetRow = ActiveCell.Row
        If targetRow < 2 Then targetRow = 2

        With ActiveSheet
            .Cells(targetRow, 7).Value = mStartTime   ' Col G: Start Time
            .Cells(targetRow, 7).NumberFormat = "hh:mm"
            .Cells(targetRow, 7).Font.Size = 11

            .Cells(targetRow, 8).Value = endTime       ' Col H: End Time
            .Cells(targetRow, 8).NumberFormat = "hh:mm"
            .Cells(targetRow, 8).Font.Size = 11

            .Cells(targetRow, 9).Formula = "=IF(AND(G" & targetRow & "<>"""",H" & targetRow & "<>""""),(H" & targetRow & "-G" & targetRow & ")*24,"""")"
            .Cells(targetRow, 9).Font.Size = 11
            .Cells(targetRow, 9).NumberFormat = "0.00"
        End With

        btn.Caption = "Start Timer"
        mStartTime = 0
    End If
End Sub
