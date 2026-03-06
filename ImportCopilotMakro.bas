Option Explicit

' Copy/Paste-fähige Einmal-Setup-Prozedur für dein Workbook.
' Fügt/aktualisiert das Modul "CopilotMakro" aus C:\temp\demosession\CopilotMakro.bas.
'
' WICHTIG (einmalig in Excel aktivieren):
' Datei > Optionen > Trust Center > Einstellungen für das Trust Center >
' Makroeinstellungen > "Zugriff auf das VBA-Projektobjektmodell vertrauen"
Public Sub ImportCopilotMakro()
    Const MODULE_PATH As String = "C:\temp\demosession\CopilotMakro.bas"
    Const MODULE_NAME As String = "CopilotMakro"

    Dim vbProj As Object
    Dim vbComp As Object

    On Error GoTo ErrHandler

    If Dir$(MODULE_PATH) = vbNullString Then
        MsgBox "Datei nicht gefunden: " & MODULE_PATH, vbExclamation
        Exit Sub
    End If

    Set vbProj = Application.VBE.ActiveVBProject

    ' Vorhandenes Modul gleichen Namens entfernen, damit Import sauber aktualisiert.
    For Each vbComp In vbProj.VBComponents
        If StrComp(vbComp.Name, MODULE_NAME, vbTextCompare) = 0 Then
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
