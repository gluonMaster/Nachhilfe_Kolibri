Attribute VB_Name = "NHblank_Sync_KinderToBlanks"
Option Explicit

' =============================================================================
' NHblank_Sync_KinderToBlanks
' -----------------------------------------------------------------------------
' Synchronizes records from Kinder sheet to Kinder_Blanks sheet.
'
' Modes:
'   - Soft sync: Updates S/T, adds new records, removes records where
'                Kinder record is inactive
'   - Hard sync: Additionally removes records from Blanks that don't exist
'                in Kinder at all
'
' Key columns: B, C, D, E, F, I, K, L, O
' Updated columns on Blanks: S, T (never G/H)
' First data row: 5
' =============================================================================

Private Const FIRST_DATA_ROW As Long = 5

' -----------------------------------------------------------------------------
' NHblank_SyncKinderToBlanks
' -----------------------------------------------------------------------------
' Main entry point for synchronization from Kinder to Kinder_Blanks.
' Prompts user to choose soft or hard sync mode.
' -----------------------------------------------------------------------------
Public Sub NHblank_SyncKinderToBlanks()
    Dim wsKinder As Worksheet
    Dim wsBlanks As Worksheet
    Dim hardSync As Boolean
    Dim msgResult As VbMsgBoxResult
    Dim appState As NHblank_AppState
    
    ' Ask user for sync mode
    msgResult = MsgBox("Moechten Sie eine harte Synchronisation ausfuehren?" & vbCrLf & vbCrLf & _
                       "Ja - Harte Synchronisation" & vbCrLf & _
                       "(loescht auch Datensaetze, die nicht in Kinder existieren)" & vbCrLf & vbCrLf & _
                       "Nein - Normale Synchronisation" & vbCrLf & _
                       "(aktualisiert nur bestehende und fuegt neue hinzu)" & vbCrLf & vbCrLf & _
                       "Abbrechen - Vorgang abbrechen", _
                       vbYesNoCancel + vbQuestion, "Synchronisation Kinder -> Kinder_Blanks")
    
    If msgResult = vbCancel Then
        Exit Sub
    End If
    
    hardSync = (msgResult = vbYes)
    
    ' Initialize application state for safe cleanup
    Set appState = New NHblank_AppState
    
    On Error GoTo ErrorHandler
    appState.Capture
    appState.OptimizeForRun
    
    ' Get worksheets
    Set wsKinder = GetKinderSheet()
    If wsKinder Is Nothing Then
        MsgBox "Blatt 'Kinder' nicht gefunden.", vbCritical, "Fehler"
        GoTo CleanUp
    End If
    
    Set wsBlanks = NHblank_EnsureKinderBlanksSheetExists()
    If wsBlanks Is Nothing Then
        MsgBox "Blatt 'Kinder_Blanks' konnte nicht erstellt/gefunden werden.", vbCritical, "Fehler"
        GoTo CleanUp
    End If
    
    ' Ensure proper formats before processing
    NHblank_EnsureBgNummerTextFormat wsKinder, FIRST_DATA_ROW
    NHblank_EnsureBgNummerTextFormat wsBlanks, FIRST_DATA_ROW
    NHblank_EnsureDateDisplayFormat wsKinder, FIRST_DATA_ROW
    NHblank_EnsureDateDisplayFormat wsBlanks, FIRST_DATA_ROW
    
    ' Execute synchronization
    Dim stats As String
    stats = ExecuteSync(wsKinder, wsBlanks, hardSync)
    
    ' Sort Kinder_Blanks by column C ascending, then renumber column A
    SortByColumnC wsBlanks
    RenumberColumnA wsBlanks
    
    ' Show results
    MsgBox "Synchronisation abgeschlossen." & vbCrLf & vbCrLf & stats, _
           vbInformation, "Synchronisation beendet"
    
CleanUp:
    appState.Restore
    Exit Sub

ErrorHandler:
    MsgBox "Fehler bei der Synchronisation:" & vbCrLf & Err.Description, _
           vbCritical, "Fehler"
    Resume CleanUp
End Sub

' -----------------------------------------------------------------------------
' ExecuteSync
' -----------------------------------------------------------------------------
' Performs the actual synchronization logic.
'
' Parameters:
'   wsKinder - Source worksheet (Kinder)
'   wsBlanks - Target worksheet (Kinder_Blanks)
'   hardSync - If True, delete records not found in Kinder
'
' Returns:
'   Statistics string for display
' -----------------------------------------------------------------------------
Private Function ExecuteSync(ByVal wsKinder As Worksheet, _
                             ByVal wsBlanks As Worksheet, _
                             ByVal hardSync As Boolean) As String
    Dim dictKinder As Object
    Dim dictBlanks As Object
    Dim dictKinderResolved As Object
    Dim dictBlanksResolved As Object
    Dim lastRowKinder As Long
    Dim lastRowBlanks As Long
    Dim addedCount As Long
    Dim updatedCount As Long
    Dim deletedCount As Long
    Dim skippedCount As Long
    
    Set dictKinder = CreateObject("Scripting.Dictionary")
    Set dictBlanks = CreateObject("Scripting.Dictionary")
    Set dictKinderResolved = CreateObject("Scripting.Dictionary")
    Set dictBlanksResolved = CreateObject("Scripting.Dictionary")
    
    ' Determine last rows
    lastRowKinder = GetLastDataRow(wsKinder)
    lastRowBlanks = GetLastDataRow(wsBlanks)
    
    ' Build dictionaries
    BuildKeyDictionary wsKinder, FIRST_DATA_ROW, lastRowKinder, dictKinder
    If lastRowBlanks >= FIRST_DATA_ROW Then
        BuildKeyDictionary wsBlanks, FIRST_DATA_ROW, lastRowBlanks, dictBlanks
    End If
    
    ' Process additions and updates
    Dim key As Variant
    Dim kinderRows As Collection
    Dim blanksRows As Collection
    Dim kinderRow As Long
    Dim blanksRow As Long
    
    For Each key In dictKinder.Keys
        Set kinderRows = dictKinder(key)
        
        ' Resolve duplicates in Kinder if needed
        kinderRow = ResolveRowCached(wsKinder, CStr(key), kinderRows, dictKinderResolved)
        If kinderRow = 0 Then
            skippedCount = skippedCount + 1
            GoTo NextKinderKey
        End If
        
        ' Check if Kinder record is active
        If Not NHblank_IsActiveRecord(wsKinder, kinderRow) Then
            ' If record exists in Blanks, mark for deletion (handled later)
            GoTo NextKinderKey
        End If
        
        ' Check if key exists in Blanks
        If dictBlanks.exists(key) Then
            ' Update existing record (only S and T)
            Set blanksRows = dictBlanks(key)
            blanksRow = ResolveRowCached(wsBlanks, CStr(key), blanksRows, dictBlanksResolved)
            If blanksRow > 0 Then
                UpdateSTColumns wsKinder, kinderRow, wsBlanks, blanksRow
                updatedCount = updatedCount + 1
            End If
        Else
            ' Add new record to Blanks
            AddRecordToBlanks wsKinder, kinderRow, wsBlanks
            addedCount = addedCount + 1
        End If
        
NextKinderKey:
    Next key
    
    ' Refresh last row after additions
    lastRowBlanks = GetLastDataRow(wsBlanks)
    
    ' Rebuild Blanks dictionary after additions
    dictBlanks.RemoveAll
    If lastRowBlanks >= FIRST_DATA_ROW Then
        BuildKeyDictionary wsBlanks, FIRST_DATA_ROW, lastRowBlanks, dictBlanks
    End If
    
    ' Collect rows to delete (process from bottom to top)
    Dim rowsToDelete As Collection
    Set rowsToDelete = New Collection
    
    Dim blanksKey As Variant
    Dim i As Long
    Dim shouldDelete As Boolean
    
    For Each blanksKey In dictBlanks.Keys
        Set blanksRows = dictBlanks(blanksKey)
        shouldDelete = False
        
        If dictKinder.exists(blanksKey) Then
            ' Key exists in Kinder - resolve once for this key
            Set kinderRows = dictKinder(blanksKey)
            kinderRow = ResolveRowCached(wsKinder, CStr(blanksKey), kinderRows, dictKinderResolved)
            
            If kinderRow = 0 Then
                ' User skipped - don't delete any rows for this key
                shouldDelete = False
            ElseIf Not NHblank_IsActiveRecord(wsKinder, kinderRow) Then
                ' Kinder record is inactive - delete all Blanks rows for this key
                shouldDelete = True
            End If
        Else
            ' Key doesn't exist in Kinder
            If hardSync Then
                shouldDelete = True
            End If
        End If
        
        ' Add all blanksRows for this key if deletion is required
        If shouldDelete Then
            For i = 1 To blanksRows.Count
                rowsToDelete.Add CLng(blanksRows(i))
            Next i
        End If
    Next blanksKey
    
    ' Delete rows from bottom to top
    deletedCount = DeleteRowsFromBottomUp(wsBlanks, rowsToDelete)
    
    ' Build statistics string
    ExecuteSync = "Hinzugefuegt: " & addedCount & vbCrLf & _
                  "Aktualisiert: " & updatedCount & vbCrLf & _
                  "Geloescht: " & deletedCount & vbCrLf & _
                  "Uebersprungen: " & skippedCount
End Function

' -----------------------------------------------------------------------------
' BuildKeyDictionary
' -----------------------------------------------------------------------------
' Builds a dictionary mapping record keys to collections of row numbers.
'
' Parameters:
'   ws        - Worksheet to scan
'   startRow  - First row to process
'   endRow    - Last row to process
'   dict      - Dictionary to populate (key -> Collection of row numbers)
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
' -----------------------------------------------------------------------------
' Returns a single row from a collection, using duplicate resolver if needed.
'
' Parameters:
'   ws   - Worksheet containing the records
'   key  - The key value (for display)
'   rows - Collection of row numbers
'
' Returns:
'   Selected row number, or 0 if skipped/cancelled
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
' -----------------------------------------------------------------------------
' Updates columns S and T on Blanks from Kinder. Does NOT touch G/H.
'
' Parameters:
'   wsSource    - Source worksheet (Kinder)
'   sourceRow   - Source row number
'   wsTarget    - Target worksheet (Kinder_Blanks)
'   targetRow   - Target row number
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
' -----------------------------------------------------------------------------
' Adds a new record from Kinder to Kinder_Blanks.
' Copies columns A through T.
'
' Parameters:
'   wsSource  - Source worksheet (Kinder)
'   sourceRow - Source row number
'   wsBlanks  - Target worksheet (Kinder_Blanks)
' -----------------------------------------------------------------------------
Private Sub AddRecordToBlanks(ByVal wsSource As Worksheet, _
                              ByVal sourceRow As Long, _
                              ByVal wsBlanks As Worksheet)
    Dim newRow As Long
    Dim col As Long
    
    ' Find next available row
    newRow = GetLastDataRow(wsBlanks) + 1
    If newRow < FIRST_DATA_ROW Then
        newRow = FIRST_DATA_ROW
    End If
    
    ' Copy columns B through T (A will be renumbered later)
    ' Skip column O (BG-Nummer) - handled separately
    For col = 2 To 20 ' B=2 to T=20
        If col <> 15 Then ' Skip column O (15)
            wsBlanks.Cells(newRow, col).value = wsSource.Cells(sourceRow, col).value
        End If
    Next col
    
    ' Write BG-Nummer (column O) using safe text conversion
    NHblank_WriteBgNummer wsBlanks.Cells(newRow, "O"), wsSource.Cells(sourceRow, "O").value
End Sub

' -----------------------------------------------------------------------------
' DeleteRowsFromBottomUp
' -----------------------------------------------------------------------------
' Deletes rows from a collection, processing from bottom to top.
'
' Parameters:
'   ws   - Worksheet to delete from
'   rows - Collection of row numbers to delete
'
' Returns:
'   Number of rows deleted
' -----------------------------------------------------------------------------
Private Function DeleteRowsFromBottomUp(ByVal ws As Worksheet, _
                                        ByVal rows As Collection) As Long
    Dim sortedRows() As Long
    Dim i As Long
    Dim j As Long
    Dim temp As Long
    Dim deletedCount As Long
    
    If rows.Count = 0 Then
        DeleteRowsFromBottomUp = 0
        Exit Function
    End If
    
    ' Copy to array for sorting
    ReDim sortedRows(1 To rows.Count)
    For i = 1 To rows.Count
        sortedRows(i) = CLng(rows(i))
    Next i
    
    ' Sort descending (bubble sort - sufficient for typical counts)
    For i = 1 To UBound(sortedRows) - 1
        For j = i + 1 To UBound(sortedRows)
            If sortedRows(i) < sortedRows(j) Then
                temp = sortedRows(i)
                sortedRows(i) = sortedRows(j)
                sortedRows(j) = temp
            End If
        Next j
    Next i
    
    ' Delete from bottom to top
    For i = 1 To UBound(sortedRows)
        ws.rows(sortedRows(i)).Delete
        deletedCount = deletedCount + 1
    Next i
    
    DeleteRowsFromBottomUp = deletedCount
End Function

' -----------------------------------------------------------------------------
' RenumberColumnA
' -----------------------------------------------------------------------------
' Renumbers column A starting from row 5 with sequential numbers.
'
' Parameters:
'   ws - Worksheet to renumber
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
' -----------------------------------------------------------------------------
' Determines the last row with data, checking both columns B and C.
'
' Parameters:
'   ws - Worksheet to check
'
' Returns:
'   Last row number with data, or FIRST_DATA_ROW - 1 if empty
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
' GetKinderSheet
' -----------------------------------------------------------------------------
' Gets the Kinder worksheet from ThisWorkbook.
'
' Returns:
'   Worksheet object, or Nothing if not found
' -----------------------------------------------------------------------------
Private Function GetKinderSheet() As Worksheet
    On Error Resume Next
    Set GetKinderSheet = ThisWorkbook.Worksheets("Kinder")
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------
' ResolveRowCached
' -----------------------------------------------------------------------------
' Returns a single row from a collection, caching the result so the user is
' not prompted more than once for the same key on the same worksheet.
' -----------------------------------------------------------------------------
Private Function ResolveRowCached(ByVal ws As Worksheet, _
                                  ByVal key As String, _
                                  ByVal rows As Collection, _
                                  ByRef cache As Object) As Long
    If cache.exists(key) Then
        ResolveRowCached = CLng(cache(key))
    Else
        Dim resolved As Long
        resolved = ResolveRowFromCollection(ws, key, rows)
        cache.Add key, resolved
        ResolveRowCached = resolved
    End If
End Function

' -----------------------------------------------------------------------------
' SortByColumnC
' -----------------------------------------------------------------------------
' Sorts data rows (from FIRST_DATA_ROW) on a worksheet by column C ascending.
' -----------------------------------------------------------------------------
Private Sub SortByColumnC(ByVal ws As Worksheet)
    Dim lastRow As Long

    lastRow = GetLastDataRow(ws)
    If lastRow < FIRST_DATA_ROW Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range("C" & FIRST_DATA_ROW & ":C" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SetRange ws.Range("A" & FIRST_DATA_ROW & ":T" & lastRow)
        .header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub
