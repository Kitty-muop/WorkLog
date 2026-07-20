"""
Inject timer button VBA into worklog.xlsm
1. Copies .xlsx -> .xlsm
2. Adds VBA TimerModule with Start/Stop logic
3. Adds Form Control button to Time Entries sheet
"""

import shutil
import os
import win32com.client as win32

XLSX_PATH = r'D:\WorkLog\worklog.xlsx'
XLSM_PATH = r'D:\WorkLog\worklog.xlsm'

VBA_CODE = '''
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
        MsgBox "Vui long dung nut nay trong sheet Time Entries.", vbExclamation, "Sai sheet"
        Exit Sub
    End If

    Set btn = ActiveSheet.Buttons(btnName)

    If btn.Caption = "Start Timer" Then
        mStartTime = Now
        btn.Caption = "Stop Timer"
        btn.Font.Bold = True
    Else
        If mStartTime = 0 Then
            MsgBox "Chua bat dau tinh gio. Nhan Start Timer truoc.", vbExclamation, "Loi"
            Exit Sub
        End If

        endTime = Now
        targetRow = ActiveCell.Row
        If targetRow < 2 Then targetRow = 2

        With ActiveSheet
            .Cells(targetRow, 7).Value = mStartTime
            .Cells(targetRow, 7).NumberFormat = "hh:mm"
            .Cells(targetRow, 7).Font.Size = 11

            .Cells(targetRow, 8).Value = endTime
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
'''


def inject_vba():
    if not os.path.exists(XLSX_PATH):
        print(f"Error: {XLSX_PATH} not found. Run build_workbook.py first.")
        return False

    shutil.copy2(XLSX_PATH, XLSM_PATH)
    print(f"Copied to {XLSM_PATH}")

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(XLSM_PATH)

        vba_module = wb.VBProject.VBComponents.Add(1)
        vba_module.Name = "TimerModule"
        vba_module.CodeModule.AddFromString(VBA_CODE)

        ws = wb.Sheets("Time Entries")

        left = ws.Range("J1").Left
        top_ = ws.Range("J1").Top
        width = 110
        height = 30

        btn = ws.Buttons.Add(left, top_, width, height)
        btn.Caption = "Start Timer"
        btn.Font.Name = "Calibri"
        btn.Font.Size = 11
        btn.Font.Bold = True
        btn.OnAction = "TimerModule.TimerButton_Click"

        wb.Save()
        print("VBA injected and button added. Saved as .xlsm")
        return True
    except Exception as e:
        print(f"Error: {e}")
        print("Hint: Check Excel Trust Center -> Macro Settings -> 'Trust access to the VBA project object model'")
        return False
    finally:
        excel.Quit()


if __name__ == '__main__':
    inject_vba()
