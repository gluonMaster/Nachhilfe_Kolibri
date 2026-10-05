Attribute VB_Name = "NHblank_Sync_BlanksToKinder"
Option Explicit

' =============================================================================
' NHblank_Sync_BlanksToKinder
' -----------------------------------------------------------------------------
' Synchronizes records from Kinder_Blanks back to Kinder.
'
' - Only processes active records from Kinder_Blanks
' - If record doesn't exist in Kinder: adds new row (copies B..T)
' - If record exists: updates ONLY G/H columns when Blanks has future dates
' - Key columns: B, C, D, E, F, I, K, L, O
' - Date columns: G (von), H (bis)
'
' Future date rule:
'   Update G/H on Kinder if H_Blanks > H_Kinder (or H_Kinder is not a date)
' =============================================================================

Private Const FIRST_DATA_ROW As Long = 5

' =============================================================================
' Private Types
' =============================================================================

Private Type SyncStats
    Added As Long
    Updated As Long
    Skipped As Long
    Inactive As Long
    DuplicateSkipped As Long
    NoUpdateNeeded As Long
End Type

' -----------------------------------------------------------------------------
' NHblank_SyncBlanksToKinder
' -----------------------------------------------------------------------------
' Public entry point: Syncs Kinder_Blanks -> Kinder
' -----------------------------------------------------------------------------
Public Sub NHblank_SyncBlanksToKinder()
    Dim wsBlanks As Worksheet
    Dim wsKinder As Worksheet
    Dim appState As NHblank_AppState
    Dim stats As SyncStats
    
    ' Initialize application state
    Set appState = New NHblank_AppState
    
    On Error GoTo ErrorHandler
    appState.Capture
    appState.OptimizeForRun
    
    ' Get worksheets
    Set wsBlanks = NHblank_GetKinderBlanksSheet()
    If wsBlanks Is Nothing Then
        MsgBox "Blatt 'Kinder_Blanks' nicht gefunden." & vbCrLf & _
               "Bitte erstellen Sie es zuerst ueber die Synchronisation Kinder -> Kinder_Blanks.", _
               vbExclamation, "Blatt fehlt"
        GoTo CleanUp
    End If
    
    Set wsKinder = NHblank_GetKinderSheet()
    If wsKinder Is Nothing Then
        MsgBox "Blatt 'Kinder' nicht gefunden.", vbCritical, "Fehler"
        GoTo CleanUp
    End If
    
    ' Apply format guards before processing
    NHblank_EnsureBgNummerTextFormat wsBlanks, FIRST_DATA_ROW
    NHblank_EnsureBgNummerTextFormat wsKinder, FIRST_DATA_ROW
    NHblank_EnsureDateDisplayFormat wsBlanks, FIRST_DATA_ROW
    NHblank_EnsureDateDisplayFormat wsKinder, FIRST_DATA_ROW
    
    ' Execute synchronization
    stats = ExecuteSync(wsBlanks, wsKinder)
    
    ' Sort Kinder by column C ascending, then renumber column A
    SortByColumnC wsKinder
    RenumberColumnA wsKinder

    ' Apply format guards after processing
    NHblank_ApplyAllFormats wsKinder, FIRST_DATA_ROW
    
    ' Show results
    ShowSyncResults stats

CleanUp:
    appState.Restore
    Exit Sub

ErrorHandler:
    MsgBox "Fehler bei der Synchronisation:" & vbCrLf & Err.Description, _
           vbCritical, "Fehler"
    Resume CleanUp
End Sub

' =============================================================================
' Private Functions
' =============================================================================

' -----------------------------------------------------------------------------
' ExecuteSync
' Performs the synchronization from Kinder_Blanks to Kinder
' -----------------------------------------------------------------------------
Private Function ExecuteSync(ByVal wsBlanks As Worksheet, _
                              ByVal wsKinder As Worksheet) As SyncStats
    Dim stats As SyncStats
    Dim dictKinder As Object
    Dim dictKinderResolved As Object
    Dim lastRowBlanks As Long
    Dim lastRowKinder As Long
    Dim blanksRow As Long
    Dim key As String
    Dim kinderRows As Collection
    Dim kinderRow As Long
    
    Set dictKinder = CreateObject("Scripting.Dictionary")
    Set dictKinderResolved = CreateObject("Scripting.Dictionary")
    
    ' Determine last rows
    lastRowBlanks = GetLastDataRow(wsBlanks)
    lastRowKinder = GetLastDataRow(wsKinder)
    
    ' Build dictionary of ALL Kinder records (active or not, for existence check)
    If lastRowKinder >= FIRST_DATA_ROW Then
        BuildKeyDictionary wsKinder, FIRST_DATA_ROW, lastRowKinder, dictKinder
    End If
    
    ' Process each row in Kinder_Blanks
    For blanksRow = FIRST_DATA_ROW To lastRowBlanks
        ' Skip empty rows
        If IsEmpty(wsBlanks.Cells(blanksRow, "C").value) Then
            stats.Skipped = stats.Skipped + 1
            GoTo NextBlanksRow
        End If
        
        ' Skip inactive records on Blanks
        If Not NHblank_IsActiveRecord(wsBlanks, blanksRow) Then
            stats.Inactive = stats.Inactive + 1
            GoTo NextBlanksRow
        End If
        
        ' Build key for this record
        key = NHblank_BuildRecordKey(wsBlanks, blanksRow)
        
        ' Check if key exists in Kinder
        If dictKinder.exists(key) Then
            ' Record exists - check if we need to update G/H
            Set kinderRows = dictKinder(key)
            kinderRow = ResolveRowCached(wsKinder, key, kinderRows, dictKinderResolved)
            
            If kinderRow = 0 Then
                ' User skipped duplicate resolution
                stats.DuplicateSkipped = stats.DuplicateSkipped + 1
                GoTo NextBlanksRow
            End If
            
            ' Check if Blanks has future dates and update if needed
            If ShouldUpdateDates(wsBlanks, blanksRow, wsKinder, kinderRow) Then
                UpdateGHColumns wsBlanks, blanksRow, wsKinder, kinderRow
                stats.Updated = stats.Updated + 1
            Else
                stats.NoUpdateNeeded = stats.NoUpdateNeeded + 1
            End If
        Else
            ' Record doesn't exist in Kinder - add new row
            AddRecordToKinder wsBlanks, blanksRow, wsKinder
            stats.Added = stats.Added + 1
            
            ' Update dictionary to handle subsequent records with same key
            Set kinderRows = New Collection
            kinderRows.Add GetLastDataRow(wsKinder)
            dictKinder.Add key, kinderRows
        End If
        
NextBlanksRow:
    Next blanksRow
    
    ExecuteSync = stats
End Function

' -----------------------------------------------------------------------------
' ShouldUpdateDates
' Determines if Blanks has "future" dates compared to Kinder
'
' Rule: Update if H_Blanks > H_Kinder (or H_Kinder is not a valid date)
'       If H_Blanks is not a date but G_Blanks is, and Kinder has no dates,
'       we update (new date info is better than none)
' -----------------------------------------------------------------------------
Private Function ShouldUpdateDates(ByVal wsBlanks As Worksheet, _
                                    ByVal blanksRow As Long, _
                                    ByVal wsKinder As Worksheet, _
                                    ByVal kinderRow As Long) As Boolean
    Dim hBlanks As Variant
    Dim hKinder As Variant
    Dim gBlanks As Variant
    Dim gKinder As Variant
    
    hBlanks = wsBlanks.Cells(blanksRow, "H").value
    hKinder = wsKinder.Cells(kinderRow, "H").value
    gBlanks = wsBlanks.Cells(blanksRow, "G").value
    gKinder = wsKinder.Cells(kinderRow, "G").value
    
    ' Case 1: H_Blanks is a valid date
    If IsDate(hBlanks) Then
        ' If H_Kinder is not a date, Blanks has more info -> update
        If Not IsDate(hKinder) Then
            ShouldUpdateDates = True
            Exit Function
        End If
        
        ' Both are dates - compare them
        If CDate(hBlanks) > CDate(hKinder) Then
            ShouldUpdateDates = True
            Exit Function
        End If
        
        ' H_Blanks <= H_Kinder, no update needed
        ShouldUpdateDates = False
        Exit Function
    End If
    
    ' Case 2: H_Blanks is not a date, check G_Blanks
    If IsDate(gBlanks) Then
        ' G_Blanks is a date, H_Blanks is not
        ' Only update if Kinder has no date info at all
        If Not IsDate(gKinder) And Not IsDate(hKinder) Then
            ShouldUpdateDates = True
            Exit Function
        End If
    End If
    
    ' Default: no update needed
    ShouldUpdateDates = False
End Function

' -----------------------------------------------------------------------------
' UpdateGHColumns
' Updates columns G and H on Kinder from Kinder_Blanks. Does NOT touch other columns.
' -----------------------------------------------------------------------------
Private Sub UpdateGHColumns(ByVal wsSource As Worksheet, _
                            ByVal sourceRow As Long, _
                            ByVal wsTarget As Worksheet, _
                            ByVal targetRow As Long)
    Dim gValue As Variant
    Dim hValue As Variant
    
    gValue = wsSource.Cells(sourceRow, "G").value
    hValue = wsSource.Cells(sourceRow, "H").value
    
    ' Write G column
    wsTarget.Cells(targetRow, "G").value = gValue
    
    ' Write H column
    wsTarget.Cells(targetRow, "H").value = hValue
    
    ' Ensure date format
    wsTarget.Cells(targetRow, "G").NumberFormat = "dd.mm.yyyy"
    wsTarget.Cells(targetRow, "H").NumberFormat = "dd.mm.yyyy"
End Sub

' -----------------------------------------------------------------------------
' AddRecordToKinder
' Adds a new record from Kinder_Blanks to Kinder. Copies columns B through T.
' Ensures BG-Nummer (column O) is stored as text.
' -----------------------------------------------------------------------------
Private Sub AddRecordToKinder(ByVal wsSource As Worksheet, _
                               ByVal sourceRow As Long, _
                               ByVal wsKinder As Worksheet)
    Dim newRow As Long
    Dim col As Long
    Dim bgValue As Variant
    
    ' Find next available row
    newRow = GetLastDataRow(wsKinder) + 1
    If newRow < FIRST_DATA_ROW Then
        newRow = FIRST_DATA_ROW
    End If
    
    ' Copy columns B through T (A will be renumbered later)
    For col = 2 To 20 ' B=2 to T=20
        ' Special handling for column O (BG-Nummer)
        If col = 15 Then
            bgValue = wsSource.Cells(sourceRow, col).value
            NHblank_WriteBgNummer wsKinder.Cells(newRow, col), bgValue
        Else
            wsKinder.Cells(newRow, col).value = wsSource.Cells(sourceRow, col).value
        End If
    Next col
    
    ' Ensure date format for G and H
    wsKinder.Cells(newRow, "G").NumberFormat = "dd.mm.yyyy"
    wsKinder.Cells(newRow, "H").NumberFormat = "dd.mm.yyyy"
End Sub

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
' RenumberColumnA
' Renumbers column A starting from FIRST_DATA_ROW with sequential numbers
' Note: We renumber the entire column to maintain consistent numbering
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
' ShowSyncResults
' Displays a summary of the sync operation (German without umlauts)
' -----------------------------------------------------------------------------
Private Sub ShowSyncResults(ByRef stats As SyncStats)
    Dim msg As String
    
    msg = "Synchronisation Kinder_Blanks -> Kinder abgeschlossen:" & vbCrLf & vbCrLf
    msg = msg & "Hinzugefuegt: " & stats.Added & vbCrLf
    msg = msg & "Aktualisiert (G/H): " & stats.Updated & vbCrLf
    
    If stats.NoUpdateNeeded > 0 Then
        msg = msg & "Keine Aktualisierung noetig: " & stats.NoUpdateNeeded & vbCrLf
    End If
    
    If stats.Inactive > 0 Then
        msg = msg & "Inaktive uebersprungen: " & stats.Inactive & vbCrLf
    End If
    
    If stats.Skipped > 0 Then
        msg = msg & "Leere Zeilen uebersprungen: " & stats.Skipped & vbCrLf
    End If
    
    If stats.DuplicateSkipped > 0 Then
        msg = msg & "Duplikate uebersprungen: " & stats.DuplicateSkipped & vbCrLf
    End If
    
    MsgBox msg, vbInformation, "Synchronisation beendet"
End Sub

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

