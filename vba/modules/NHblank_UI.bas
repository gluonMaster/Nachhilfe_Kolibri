Attribute VB_Name = "NHblank_UI"
Option Explicit

' Ask user whether to process selected records or all active records
' Returns True if user made a choice, False if cancelled
Public Function AskProcessingMode(ByRef processAllRecords As Boolean) As Boolean
    Dim response As VbMsgBoxResult
    
    response = MsgBox( _
        "Moechten Sie Blanks nur fuer ausgewaehlte Zeilen erstellen?" & vbCrLf & vbCrLf & _
        "Ja - Nur ausgewaehlte Zeilen auf 'Kinder_Blanks'" & vbCrLf & _
        "Nein - Alle aktiven Datensaetze", _
        vbYesNoCancel + vbQuestion, _
        "Modus auswaehlen" _
    )
    
    Select Case response
        Case vbYes
            processAllRecords = False
            AskProcessingMode = True
        Case vbNo
            processAllRecords = True
            AskProcessingMode = True
        Case vbCancel
            AskProcessingMode = False
    End Select
End Function

' Let user select target folder for saving blanks
' Returns folder path or empty string if cancelled
Public Function SelectTargetFolder() As String
    Dim folderDialog As fileDialog
    Dim defaultFolder As String
    
    Set folderDialog = Application.fileDialog(msoFileDialogFolderPicker)
    defaultFolder = GetDefaultTargetFolder()
    
    With folderDialog
        .title = "Zielordner fuer Blanks auswaehlen"
        .AllowMultiSelect = False
        .InitialFileName = defaultFolder & "\"
        
        If .Show = -1 Then
            SelectTargetFolder = .SelectedItems(1)
        Else
            SelectTargetFolder = ""
        End If
    End With
    
    Set folderDialog = Nothing
End Function

' Let user select the Word template (Shablon.docx).
Public Function SelectWordTemplateFile() As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .title = "Bitte waehlen Sie die Datei Shablon.docx aus"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Word Dateien", "*.docx"
        .Filters.Add "Alle Dateien", "*.*"
        
        ' Try to set initial directory to user's Documents folder
        .InitialFileName = Environ("USERPROFILE") & "\Documents\"
        
        If .Show = -1 Then
            SelectWordTemplateFile = .SelectedItems(1)
        Else
            SelectWordTemplateFile = ""
        End If
    End With
    
    Set fileDialog = Nothing
End Function

' Let user select the legacy Sozialamt template (Shablon.xlsx).
Public Function SelectLegacyTemplateFile() As String
    Dim fileDialog As fileDialog

    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)

    With fileDialog
        .title = "Bitte waehlen Sie die Datei Shablon.xlsx aus"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Dateien", "*.xlsx;*.xls"
        .Filters.Add "Alle Dateien", "*.*"
        .InitialFileName = Environ("USERPROFILE") & "\Documents\"

        If .Show = -1 Then
            SelectLegacyTemplateFile = .SelectedItems(1)
        Else
            SelectLegacyTemplateFile = ""
        End If
    End With

    Set fileDialog = Nothing
End Function

' Backward-compatible alias for older calls.
Public Function SelectTemplateFile() As String
    SelectTemplateFile = SelectWordTemplateFile()
End Function

' Ask once how file-name collisions should be handled for this run.
Public Function AskOverwriteMode(ByRef overwriteExisting As Boolean) As Boolean
    Dim response As VbMsgBoxResult

    response = MsgBox( _
        "Sollen bereits vorhandene gleichnamige Dateien ueberschrieben werden?" & _
        vbCrLf & vbCrLf & _
        "Ja - Vorhandene Dateien sicher ersetzen" & vbCrLf & _
        "Nein - Vorhandene Dateien ueberspringen" & vbCrLf & _
        "Abbrechen - Keine Blanks erstellen", _
        vbYesNoCancel + vbQuestion, _
        "Vorhandene Dateien" _
    )

    Select Case response
        Case vbYes
            overwriteExisting = True
            AskOverwriteMode = True
        Case vbNo
            overwriteExisting = False
            AskOverwriteMode = True
        Case vbCancel
            AskOverwriteMode = False
    End Select
End Function

' Show error message to user
Public Sub ShowError(ByVal message As String)
    MsgBox message, vbCritical, "Fehler"
End Sub

' Show warning about wrong worksheet
Public Sub ShowWrongSheetWarning()
    MsgBox "Bitte waehlen Sie Zeilen auf dem Blatt 'Kinder_Blanks' aus.", vbExclamation, "Falsches Blatt"
End Sub

' Show success message with count of created files and optional warnings.
Public Sub ShowSuccess(ByVal filesCreated As Long, _
                       Optional ByVal warnings As String = "")
    Dim message As String
    Dim messageStyle As VbMsgBoxStyle

    message = "Erfolgreich abgeschlossen!" & vbCrLf & vbCrLf & _
              "Anzahl erstellter Dateien: " & CStr(filesCreated)
    messageStyle = vbInformation

    If Len(Trim$(warnings)) > 0 Then
        message = message & vbCrLf & vbCrLf & _
                  "Warnungen / uebersprungene Dateien:" & vbCrLf & _
                  warnings
        messageStyle = vbExclamation
    End If

    MsgBox message, messageStyle, "Fertig"
End Sub

' Returns default output folder on Desktop and creates it if missing
Private Function GetDefaultTargetFolder() As String
    Dim desktopPath As String
    Dim defaultFolder As String
    
    desktopPath = Environ$("USERPROFILE") & "\Desktop"
    defaultFolder = desktopPath & "\Nachhilfe Formular"
    
    On Error Resume Next
    If Dir(defaultFolder, vbDirectory) = "" Then
        MkDir defaultFolder
    End If
    On Error GoTo 0
    
    If Dir(defaultFolder, vbDirectory) <> "" Then
        GetDefaultTargetFolder = defaultFolder
    Else
        GetDefaultTargetFolder = desktopPath
    End If
End Function

