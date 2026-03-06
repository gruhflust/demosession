Attribute VB_Name = "CopilotMakro"
Option Explicit

Public Sub CopilotMakro()
    Dim ws As Worksheet
    Dim wb As Workbook

    Set wb = ActiveWorkbook

    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("Copilot").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Copilot"
    ws.Range("A1").Value = "Hallo von Copilot"
End Sub
