Attribute VB_Name = "NHblank_Main"
Option Explicit

' Main entry point for the blank generation macro
Public Sub GenerateBlanks()
    Dim screenUpdateState As Boolean
    Dim calculationState As XlCalculation
    Dim eventsState As Boolean
    Dim selectedRows As Collection
    Dim selectionSnapshot As Range
    Dim wsKinderBlanks As Worksheet
    
    On Error GoTo ErrorHandler
    
    ' Freeze the exact selection before any dialogs can change Excel state.
    If TypeName(selection) = "Range" Then
        Set selectionSnapshot = selection
    End If
    
    ' Save current Excel state
    screenUpdateState = Application.ScreenUpdating
    calculationState = Application.Calculation
    eventsState = Application.EnableEvents
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Pre-flight diagnostic check (optional - comment out in production)
    ' Uncomment next line to run diagnostics before each execution:
    ' NHblank_Diagnostics.RunFullDiagnostics
    
    ' Ask user for processing mode
    Dim processAllRecords As Boolean
    If Not NHblank_UI.AskProcessingMode(processAllRecords) Then
        ' User cancelled
        GoTo CleanExit
    End If
    
    ' In selection mode, validate that rows are selected on Kinder_Blanks
    If Not processAllRecords Then
        Set wsKinderBlanks = NHblank_Sheets.NHblank_GetKinderBlanksSheet()
        
        If wsKinderBlanks Is Nothing Then
            NHblank_UI.ShowError "Blatt 'Kinder_Blanks' nicht gefunden." & vbCrLf & _
                                 "Bitte zuerst die Synchronisation ausfuehren."
            GoTo CleanExit
        End If
        
        If selectionSnapshot Is Nothing Then
            NHblank_UI.ShowError "Keine Zellenauswahl gefunden." & vbCrLf & _
                                 "Bitte waehlen Sie die gewuenschten Zeilen erneut aus."
            GoTo CleanExit
        End If

        If Not selectionSnapshot.Worksheet Is wsKinderBlanks Then
            NHblank_UI.ShowWrongSheetWarning
            GoTo CleanExit
        End If
        
        Set selectedRows = CollectSelectedRows(selectionSnapshot, _
                                               wsKinderBlanks, 5)
        If selectedRows.Count = 0 Then
            NHblank_UI.ShowError "Keine gueltigen Zeilen ausgewaehlt." & vbCrLf & _
                                 "Bitte waehlen Sie Zeilen ab Zeile 5 auf 'Kinder_Blanks' aus."
            GoTo CleanExit
        End If
    End If
    
    ' Ask user to select target folder
    Dim targetFolder As String
    targetFolder = NHblank_UI.SelectTargetFolder()
    If targetFolder = "" Then
        ' User cancelled folder selection
        GoTo CleanExit
    End If
    
    ' Resolve Word and legacy Excel templates.
    Dim wordTemplatePath As String
    Dim legacyTemplatePath As String
    Dim workbookLocalPath As String
    Dim errorMsg As String
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get local path (handles OneDrive URLs)
    workbookLocalPath = NHblank_Utils.GetLocalPath(ThisWorkbook.Path)
    
    If workbookLocalPath = "" Then
        ' Could not determine local path, ask for both templates manually.
        MsgBox "Automatische Erkennung des Template-Pfades fehlgeschlagen." & vbCrLf & _
               "Bitte waehlen Sie Shablon.docx und Shablon.xlsx manuell aus.", _
               vbInformation
        
        wordTemplatePath = NHblank_UI.SelectWordTemplateFile()
        If wordTemplatePath = "" Then GoTo CleanExit

        legacyTemplatePath = NHblank_UI.SelectLegacyTemplateFile()
        If legacyTemplatePath = "" Then GoTo CleanExit
    Else
        wordTemplatePath = fso.BuildPath(workbookLocalPath, "Shablon.docx")
        legacyTemplatePath = fso.BuildPath(workbookLocalPath, "Shablon.xlsx")

        If Not fso.FileExists(wordTemplatePath) Then
            MsgBox "Datei Shablon.docx wurde im Verzeichnis nicht gefunden: " & _
                   workbookLocalPath & vbCrLf & _
                   "Bitte waehlen Sie die Datei manuell aus.", vbExclamation

            wordTemplatePath = NHblank_UI.SelectWordTemplateFile()
            If wordTemplatePath = "" Then GoTo CleanExit
        End If

        If Not fso.FileExists(legacyTemplatePath) Then
            MsgBox "Datei Shablon.xlsx wurde im Verzeichnis nicht gefunden: " & _
                   workbookLocalPath & vbCrLf & _
                   "Bitte waehlen Sie die Datei manuell aus.", vbExclamation

            legacyTemplatePath = NHblank_UI.SelectLegacyTemplateFile()
            If legacyTemplatePath = "" Then GoTo CleanExit
        End If
    End If
    
    ' Pre-process validation
    If Not NHblank_Diagnostics.PreProcessCheck( _
        wordTemplatePath, legacyTemplatePath, targetFolder, errorMsg) Then
        NHblank_UI.ShowError "Validierungsfehler: " & errorMsg
        GoTo CleanExit
    End If

    ' Ask once whether existing files may be replaced in this run.
    Dim overwriteExisting As Boolean
    If Not NHblank_UI.AskOverwriteMode(overwriteExisting) Then
        GoTo CleanExit
    End If
    
    ' Process records
    Dim blanksCreated As Long
    Dim activeRecordsFound As Long
    Dim processWarnings As String
    blanksCreated = NHblank_DataProcessor.ProcessRecords( _
        processAllRecords, _
        wordTemplatePath, _
        legacyTemplatePath, _
        targetFolder, _
        activeRecordsFound, _
        selectedRows, _
        overwriteExisting, _
        processWarnings _
    )
    
    ' Show appropriate message based on results
    If blanksCreated > 0 Then
        NHblank_UI.ShowSuccess blanksCreated, processWarnings
    ElseIf activeRecordsFound = 0 Then
        NHblank_UI.ShowError "Keine aktiven Datensaetze gefunden." & vbCrLf & vbCrLf & _
                             "Hinweis: Aktive Datensaetze muessen:" & vbCrLf & _
                             "- Nicht leere Zellen in Spalte C haben" & vbCrLf & _
                             "- NICHT grauen Text haben (ColorIndex <> 15)"
    Else
        NHblank_UI.ShowError _
            "Es wurden " & activeRecordsFound & _
            " aktive Datensaetze gefunden, aber keine Datei wurde erstellt." & _
            vbCrLf & vbCrLf & processWarnings
    End If
    
CleanExit:
    ' Restore Excel state
    Application.ScreenUpdating = screenUpdateState
    Application.Calculation = calculationState
    Application.EnableEvents = eventsState
    Exit Sub
    
ErrorHandler:
    ' Restore Excel state even on error
    Application.ScreenUpdating = screenUpdateState
    Application.Calculation = calculationState
    Application.EnableEvents = eventsState
    
    NHblank_UI.ShowError "Fehler beim Ausfuehren des Makros: " & Err.Description
End Sub

' Temporary debugging procedure - call this to see path information
Public Sub DebugPaths()
    NHblank_Utils.ShowPathDebugInfo
End Sub

' Run full system diagnostics
Public Sub RunFullDiagnostics()
    NHblank_Diagnostics.RunFullDiagnostics
End Sub

' Collect unique visible row numbers from an immutable selection snapshot.
Private Function CollectSelectedRows(ByVal selectedRange As Range, _
                                     ByVal ws As Worksheet, _
                                     ByVal firstDataRow As Long) As Collection
    Dim result As Collection
    Dim seenRows As Object
    Dim area As Range
    Dim selectedRow As Range
    Dim rowNum As Long
    
    Set result = New Collection
    Set seenRows = CreateObject("Scripting.Dictionary")
    
    If selectedRange Is Nothing Or ws Is Nothing Then
        Set CollectSelectedRows = result
        Exit Function
    End If
    
    If Not selectedRange.Worksheet Is ws Then
        Set CollectSelectedRows = result
        Exit Function
    End If
    
    For Each area In selectedRange.Areas
        For Each selectedRow In area.rows
            rowNum = selectedRow.row
            If rowNum >= firstDataRow And _
               Not selectedRow.EntireRow.Hidden Then
                If Not seenRows.exists(rowNum) Then
                    seenRows.Add rowNum, True
                    result.Add rowNum
                End If
            End If
        Next selectedRow
    Next area
    
    Set CollectSelectedRows = result
End Function

