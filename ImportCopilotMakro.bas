Attribute VB_Name = "ImportCopilotMakro"
Option Explicit

' Importiert die von Copilot gelieferte .bas-Datei in das aktive VBA-Projekt.
' Hinweis: In Excel muss aktiviert sein:
' Trust Center -> Makroeinstellungen -> "Zugriff auf das VBA-Projektobjektmodell vertrauen".
Public Sub ImportCopilotMakro()
    Const MODULE_PATH As String = "C:\temp\demosession\CopilotMakro.bas"
    Const MODULE_NAME As String = "CopilotMakro"

    Dim vbProj As Object
    Dim vbComp As Object

    On Error GoTo ErrHandler

    Set vbProj = Application.VBE.ActiveVBProject

    For Each vbComp In vbProj.VBComponents
        If vbComp.Name = MODULE_NAME Then
            vbProj.VBComponents.Remove vbComp
            Exit For
        End If
    Next vbComp

    vbProj.VBComponents.Import MODULE_PATH

    MsgBox "Import erfolgreich: " & MODULE_NAME, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Import fehlgeschlagen: " & Err.Description, vbCritical
End Sub
