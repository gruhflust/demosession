Attribute VB_Name = "CopilotMakro"
Option Explicit

Public Sub CopilotMakro()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim monate As Variant
    Dim i As Integer

    Set wb = ActiveWorkbook

    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("Copilot").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Copilot"
    ws.Range("A1").Value = "Hallo von Copilot"
    
    ' Reiter in Blau färben
    ws.Tab.Color = RGB(0, 0, 255)
    
    ' Monatsnamen in A2:A13 eintragen
    monate = Array("Januar", "Februar", "M" & Chr(228) & "rz", "April", "Mai", "Juni", _
                   "Juli", "August", "September", "Oktober", "November", "Dezember")
    
    For i = 0 To 11
        ws.Range("A" & (i + 2)).Value = monate(i)
    Next i
End Sub
