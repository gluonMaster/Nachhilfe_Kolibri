Attribute VB_Name = "NHblank_TransferSelection"
Option Explicit

' =============================================================================
' NHblank_TransferSelection
' -----------------------------------------------------------------------------
' Transfers selected rows from Kinder sheet to Kinder_Blanks.
' Supports non-contiguous (multi-area) selection.
'
' - Only works when active sheet is "Kinder"
' - Processes rows >= 5 only
' - Inactive records are skipped (with optional confirmation)
' - New records are added to Kinder_Blanks
' - Existing records (by key) are updated (S/T only, not G/H)
' - Duplicates are resolved via NHblank_ResolveDuplicateRow
' - BG-Nummer format is protected via NHblank_FormatGuard
' =============================================================================

Private Const FIRST_DATA_ROW As Long = 5
Private Const SHEET_KINDER As String = "Kinder"

' =============================================================================
' Private Types
' =============================================================================

Private Type TransferStats
    Added As Long
    Updated As Long
    Skipped As Long
    Inactive As Long
    DuplicateSkipped As Long
End Type

' -----------------------------------------------------------------------------
' NHblank_TransferSelectedKinderRowsToBlanks
' -----------------------------------------------------------------------------
' Public entry point: Transfers selected rows from Kinder to Kinder_Blanks.
' -----------------------------------------------------------------------------
Public Sub NHblank_TransferSelectedKinderRowsToBlanks()
    Dim wsKinder As Worksheet
    Dim wsBlanks As Worksheet
    Dim appState As NHblank_AppState
    Dim selectedRows As Collection
    Dim selectionSnapshot As Range
    Dim stats As TransferStats
    
    ' Verify active sheet is Kinder
    If Not IsActiveSheetKinder() Then
        MsgBox "Diese Funktion ist nur auf dem Blatt 'Kinder' verfuegbar." & vbCrLf & _
               "Bitte wechseln Sie zum Blatt 'Kinder' und versuchen Sie es erneut.", _
               vbExclamation, "Falsches Blatt"
        Exit Sub
    End If
    
    ' Get Kinder sheet
    Set wsKinder = ActiveSheet

    If TypeName(selection) = "Range" Then
        Set selectionSnapshot = selection
    End If
    
    ' Collect selected row numbers (handles non-contiguous selection)
    Set selectedRows = CollectSelectedRows(selectionSnapshot, wsKinder)
    
    If selectedRows.Count = 0 Then
        MsgBox "Keine gueltigen Zeilen ausgewaehlt." & vbCrLf & _
               "Bitte waehlen Sie Zeilen ab Zeile " & FIRST_DATA_ROW & " aus.", _
               vbExclamation, "Keine Auswahl"
        Exit Sub
    End If
    
    ' Initialize application state for optimization
    Set appState = New NHblank_AppState
    
    On Error GoTo ErrorHandler
    appState.Capture
    appState.OptimizeForRun
    
    ' Ensure Kinder_Blanks sheet exists
    Set wsBlanks = NHblank_EnsureKinderBlanksSheetExists()
    If wsBlanks Is Nothing Then
        MsgBox "Blatt 'Kinder_Blanks' konnte nicht erstellt werden.", _
               vbCritical, "Fehler"
        GoTo CleanUp
    End If
    
    ' Apply format guards before processing
    NHblank_EnsureBgNummerTextFormat wsKinder, FIRST_DATA_ROW
    NHblank_EnsureBgNummerTextFormat wsBlanks, FIRST_DATA_ROW
    NHblank_EnsureDateDisplayFormat wsKinder, FIRST_DATA_ROW
    NHblank_EnsureDateDisplayFormat wsBlanks, FIRST_DATA_ROW
    
    ' Process selected rows
    stats = ProcessSelectedRows(wsKinder, wsBlanks, selectedRows)
    
    ' Renumber column A on Kinder_Blanks
    RenumberColumnA wsBlanks
    
    ' Apply format guards after processing
    NHblank_ApplyAllFormats wsBlanks, FIRST_DATA_ROW
    
    ' Show results
    ShowTransferResults stats
    
CleanUp:
    appState.Restore
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Uebertragen der Datensaetze:" & vbCrLf & Err.Description, _
           vbCritical, "Fehler"
    Resume CleanUp
End Sub

' =============================================================================
' Private Functions
' =============================================================================

' -----------------------------------------------------------------------------
' IsActiveSheetKinder
' Returns True if the active sheet is named "Kinder"
' -----------------------------------------------------------------------------
Private Function IsActiveSheetKinder() As Boolean
    On Error Resume Next
    IsActiveSheetKinder = (ActiveSheet.name = SHEET_KINDER)
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------
' CollectSelectedRows
' Collects all selected row numbers from Selection.Areas (handles non-contiguous)
' Only includes rows >= FIRST_DATA_ROW
' Returns a Collection of unique row numbers (Long)
' -----------------------------------------------------------------------------
Private Function CollectSelectedRows(ByVal selectedRange As Range, _
                                     ByVal ws As Worksheet) As Collection
    Dim result As Collection
    Dim area As Range
    Dim row As Range
    Dim rowNum As Long
    Dim dictSeen As Object
    
    Set result = New Collection
    Set dictSeen = CreateObject("Scripting.Dictionary")
    
    If selectedRange Is Nothing Or ws Is Nothing Then
        Set CollectSelectedRows = result
        Exit Function
    End If

    If Not selectedRange.Worksheet Is ws Then
        Set CollectSelectedRows = result
        Exit Function
    End If
    
    ' Iterate through all selection areas (for non-contiguous selection)
    For Each area In selectedRange.Areas
        ' Iterate through each row in the area
        For Each row In area.rows
            rowNum = row.row
            
            ' Only include visible data rows (>= 5) and avoid duplicates.
            If rowNum >= FIRST_DATA_ROW And _
               Not row.EntireRow.Hidden Then
                If Not dictSeen.exists(rowNum) Then
                    dictSeen.Add rowNum, True
                    result.Add rowNum
                End If
            End If
        Next row
    Next area
    
    Set CollectSelectedRows = result
End Function

' -----------------------------------------------------------------------------
' ProcessSelectedRows
' Processes each selected row: adds new or updates existing records
' Returns statistics about the operation
' -----------------------------------------------------------------------------
Private Function ProcessSelectedRows(ByVal wsKinder As Worksheet, _
                                      ByVal wsBlanks As Worksheet, _
                                      ByVal selectedRows As Collection) As TransferStats
    Dim stats As TransferStats
    Dim dictBlanks As Object
    Dim i As Long
    Dim kinderRow As Long
    Dim key As String
    Dim blanksRows As Collection
    Dim blanksRow As Long
    Dim lastRowBlanks As Long
    Dim msgResult As VbMsgBoxResult
    
    Set dictBlanks = CreateObject("Scripting.Dictionary")
    
    ' Build dictionary of existing keys on Kinder_Blanks
    lastRowBlanks = GetLastDataRow(wsBlanks)
    If lastRowBlanks >= FIRST_DATA_ROW Then
        BuildKeyDictionary wsBlanks, FIRST_DATA_ROW, lastRowBlanks, dictBlanks
    End If
    
    ' Process each selected row
    For i = 1 To selectedRows.Count
        kinderRow = CLng(selectedRows(i))
        
        ' Skip empty rows (no data in column C)
        If IsEmpty(wsKinder.Cells(kinderRow, "C").value) Then
            stats.Skipped = stats.Skipped + 1
            GoTo NextRow
        End If
        
        ' Check if record is active
        If Not NHblank_IsActiveRecord(wsKinder, kinderRow) Then
            ' Ask user whether to include inactive record
            msgResult = MsgBox("Zeile " & kinderRow & " ist inaktiv (grau markiert)." & vbCrLf & vbCrLf & _
                              "Moechten Sie diesen Datensatz trotzdem uebertragen?", _
                              vbYesNo + vbQuestion, "Inaktiver Datensatz")
            
            If msgResult = vbNo Then
                stats.Inactive = stats.Inactive + 1
                GoTo NextRow
            End If
        End If
        
        ' Build key for this record
        key = NHblank_BuildRecordKey(wsKinder, kinderRow)
        
        ' Check if key exists in Blanks
        If dictBlanks.exists(key) Then
            ' Key exists - update existing record (only S/T)
            Set blanksRows = dictBlanks(key)
            blanksRow = ResolveRowFromCollection(wsBlanks, key, blanksRows)
            
            If blanksRow = 0 Then
                ' User skipped duplicate resolution
                stats.DuplicateSkipped = stats.DuplicateSkipped + 1
                GoTo NextRow
            End If
            
            ' Update S and T columns only
            UpdateSTColumns wsKinder, kinderRow, wsBlanks, blanksRow
            stats.Updated = stats.Updated + 1
        Else
            ' Key doesn't exist - add new record
            AddRecordToBlanks wsKinder, kinderRow, wsBlanks
            stats.Added = stats.Added + 1
            
            ' Add to dictionary to handle subsequent duplicates in selection
            Set blanksRows = New Collection
            blanksRows.Add GetLastDataRow(wsBlanks)
            dictBlanks.Add key, blanksRows
        End If
        
NextRow:
    Next i
    
    ProcessSelectedRows = stats
End Function

' -----------------------------------------------------------------------------
' BuildKeyDictionary
' Builds a dictionary mapping record keys to collections of row numbers
' -----------------------------------------------------------------------------
Private Sub BuildKeyDictionary(ByVal ws As Worksheet, _
                               ByVal startRow As Long, _
                               ByVal endRow As Long, _
                               ByRef dict As Object)
    Dim rowNum As Long
    Dim key As String
    Dim rows As Collection
    
    For rowNum = startRow To endRow
        ' Skip completely empty rows
        If Not IsEmpty(ws.Cells(rowNum, "C").value) Then
            key = NHblank_BuildRecordKey(ws, rowNum)
            
            If dict.exists(key) Then
                Set rows = dict(key)
                rows.Add rowNum
            Else
                Set rows = New Collection
                rows.Add rowNum
                dict.Add key, rows
            End If
        End If
    Next rowNum
End Sub

' -----------------------------------------------------------------------------
' ResolveRowFromCollection
' Returns a single row from a collection, using duplicate resolver if needed
' -----------------------------------------------------------------------------
Private Function ResolveRowFromCollection(ByVal ws As Worksheet, _
                                          ByVal key As String, _
                                          ByVal rows As Collection) As Long
    If rows.Count = 0 Then
        ResolveRowFromCollection = 0
    ElseIf rows.Count = 1 Then
        ResolveRowFromCollection = CLng(rows(1))
    Else
        ResolveRowFromCollection = NHblank_ResolveDuplicateRow(ws, key, rows)
    End If
End Function

' -----------------------------------------------------------------------------
' UpdateSTColumns
' Updates columns S and T on Blanks from Kinder. Does NOT touch G/H.
' -----------------------------------------------------------------------------
Private Sub UpdateSTColumns(ByVal wsSource As Worksheet, _
                            ByVal sourceRow As Long, _
                            ByVal wsTarget As Worksheet, _
                            ByVal targetRow As Long)
    ' Copy S column
    wsTarget.Cells(targetRow, "S").value = wsSource.Cells(sourceRow, "S").value
    
    ' Copy T column
    wsTarget.Cells(targetRow, "T").value = wsSource.Cells(sourceRow, "T").value
End Sub

' -----------------------------------------------------------------------------
' AddRecordToBlanks
' Adds a new record from Kinder to Kinder_Blanks. Copies columns B through T.
' Ensures BG-Nummer (column O) is stored as text.
' -----------------------------------------------------------------------------
Private Sub AddRecordToBlanks(ByVal wsSource As Worksheet, _
                              ByVal sourceRow As Long, _
                              ByVal wsBlanks As Worksheet)
    Dim newRow As Long
    Dim col As Long
    Dim bgValue As Variant
    
    ' Find next available row
    newRow = GetLastDataRow(wsBlanks) + 1
    If newRow < FIRST_DATA_ROW Then
        newRow = FIRST_DATA_ROW
    End If
    
    ' Copy columns B through T (A will be renumbered later)
    For col = 2 To 20 ' B=2 to T=20
        ' Special handling for column O (BG-Nummer)
        If col = 15 Then
            bgValue = wsSource.Cells(sourceRow, col).value
            NHblank_WriteBgNummer wsBlanks.Cells(newRow, col), bgValue
        Else
            wsBlanks.Cells(newRow, col).value = wsSource.Cells(sourceRow, col).value
        End If
    Next col
End Sub

' -----------------------------------------------------------------------------
' RenumberColumnA
' Renumbers column A starting from FIRST_DATA_ROW with sequential numbers
' -----------------------------------------------------------------------------
Private Sub RenumberColumnA(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim rowNum As Long
    Dim counter As Long
    
    lastRow = GetLastDataRow(ws)
    If lastRow < FIRST_DATA_ROW Then
        Exit Sub
    End If
    
    counter = 1
    For rowNum = FIRST_DATA_ROW To lastRow
        ' Only number rows that have data in column C
        If Not IsEmpty(ws.Cells(rowNum, "C").value) Then
            ws.Cells(rowNum, "A").value = counter
            counter = counter + 1
        Else
            ws.Cells(rowNum, "A").value = ""
        End If
    Next rowNum
End Sub

' -----------------------------------------------------------------------------
' GetLastDataRow
' Determines the last row with data, checking both columns B and C
' -----------------------------------------------------------------------------
Private Function GetLastDataRow(ByVal ws As Worksheet) As Long
    Dim lastRowB As Long
    Dim lastRowC As Long
    
    On Error Resume Next
    lastRowB = ws.Cells(ws.rows.Count, "B").End(xlUp).row
    lastRowC = ws.Cells(ws.rows.Count, "C").End(xlUp).row
    On Error GoTo 0
    
    GetLastDataRow = lastRowB
    If lastRowC > GetLastDataRow Then
        GetLastDataRow = lastRowC
    End If
    
    ' Ensure we don't go below header row
    If GetLastDataRow < FIRST_DATA_ROW Then
        GetLastDataRow = FIRST_DATA_ROW - 1
    End If
End Function

' -----------------------------------------------------------------------------
' ShowTransferResults
' Displays a summary of the transfer operation
' -----------------------------------------------------------------------------
Private Sub ShowTransferResults(ByRef stats As TransferStats)
    Dim msg As String
    
    msg = "Uebertragung abgeschlossen:" & vbCrLf & vbCrLf
    msg = msg & "Hinzugefuegt: " & stats.Added & vbCrLf
    msg = msg & "Aktualisiert (S/T): " & stats.Updated & vbCrLf
    
    If stats.Inactive > 0 Then
        msg = msg & "Inaktive uebersprungen: " & stats.Inactive & vbCrLf
    End If
    
    If stats.Skipped > 0 Then
        msg = msg & "Leere Zeilen uebersprungen: " & stats.Skipped & vbCrLf
    End If
    
    If stats.DuplicateSkipped > 0 Then
        msg = msg & "Duplikate uebersprungen: " & stats.DuplicateSkipped & vbCrLf
    End If
    
    MsgBox msg, vbInformation, "Uebertragung beendet"
End Sub

