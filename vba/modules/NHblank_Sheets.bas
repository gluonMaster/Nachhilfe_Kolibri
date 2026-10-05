Attribute VB_Name = "NHblank_Sheets"
Option Explicit

' =============================================================================
' NHblank_Sheets
' Utilities for safely accessing and creating worksheets.
' Handles creation and setup of Kinder_Blanks sheet.
' =============================================================================

' Constants
Private Const SHEET_KINDER As String = "Kinder"
Private Const SHEET_KINDER_BLANKS As String = "Kinder_Blanks"
Private Const FIRST_DATA_ROW As Long = 5
Private Const HEADER_LAST_ROW As Long = 4
Private Const LAST_STRUCTURE_COL As String = "T"

' -----------------------------------------------------------------------------
' NHblank_TryGetWorksheet
' Safely attempts to get a worksheet by name from the specified workbook.
' Returns True if found, False otherwise. Output parameter ws is set or Nothing.
' -----------------------------------------------------------------------------
Public Function NHblank_TryGetWorksheet(ByVal wb As Workbook, _
                                         ByVal sheetName As String, _
                                         ByRef ws As Worksheet) As Boolean
    Set ws = Nothing
    NHblank_TryGetWorksheet = False
    
    If wb Is Nothing Then Exit Function
    If Len(Trim(sheetName)) = 0 Then Exit Function
    
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    NHblank_TryGetWorksheet = Not (ws Is Nothing)
End Function

' -----------------------------------------------------------------------------
' NHblank_GetKinderSheet
' Returns the Kinder worksheet from ThisWorkbook, or Nothing if not found.
' -----------------------------------------------------------------------------
Public Function NHblank_GetKinderSheet() As Worksheet
    Dim ws As Worksheet
    If NHblank_TryGetWorksheet(ThisWorkbook, SHEET_KINDER, ws) Then
        Set NHblank_GetKinderSheet = ws
    Else
        Set NHblank_GetKinderSheet = Nothing
    End If
End Function

' -----------------------------------------------------------------------------
' NHblank_GetKinderBlanksSheet
' Returns the Kinder_Blanks worksheet from ThisWorkbook, or Nothing if not found.
' -----------------------------------------------------------------------------
Public Function NHblank_GetKinderBlanksSheet() As Worksheet
    Dim ws As Worksheet
    If NHblank_TryGetWorksheet(ThisWorkbook, SHEET_KINDER_BLANKS, ws) Then
        Set NHblank_GetKinderBlanksSheet = ws
    Else
        Set NHblank_GetKinderBlanksSheet = Nothing
    End If
End Function

' -----------------------------------------------------------------------------
' NHblank_EnsureKinderBlanksSheetExists
' Returns existing Kinder_Blanks sheet or creates it from Kinder template.
' Applies format guards for BG-Nummer (column O) and dates (G/H, T2).
' -----------------------------------------------------------------------------
Public Function NHblank_EnsureKinderBlanksSheetExists() As Worksheet
    Dim wsBlanks As Worksheet
    Dim wsKinder As Worksheet
    Dim rngHeader As Range
    Dim lastCol As Long
    
    ' Check if Kinder_Blanks already exists
    If NHblank_TryGetWorksheet(ThisWorkbook, SHEET_KINDER_BLANKS, wsBlanks) Then
        ' Apply format guards to existing sheet
        NHblank_EnsureBgNummerTextFormat wsBlanks, FIRST_DATA_ROW
        NHblank_EnsureDateDisplayFormat wsBlanks, FIRST_DATA_ROW
        Set NHblank_EnsureKinderBlanksSheetExists = wsBlanks
        Exit Function
    End If
    
    ' Get source Kinder sheet
    If Not NHblank_TryGetWorksheet(ThisWorkbook, SHEET_KINDER, wsKinder) Then
        ' Cannot create without source - return Nothing
        Set NHblank_EnsureKinderBlanksSheetExists = Nothing
        Exit Function
    End If
    
    ' Create new Kinder_Blanks sheet
    Set wsBlanks = CreateKinderBlanksSheet(wsKinder)
    
    If Not wsBlanks Is Nothing Then
        ' Apply format guards
        NHblank_EnsureBgNummerTextFormat wsBlanks, FIRST_DATA_ROW
        NHblank_EnsureDateDisplayFormat wsBlanks, FIRST_DATA_ROW
    End If
    
    Set NHblank_EnsureKinderBlanksSheetExists = wsBlanks
End Function

' =============================================================================
' Private Helpers
' =============================================================================

' -----------------------------------------------------------------------------
' CreateKinderBlanksSheet
' Creates a new Kinder_Blanks sheet by copying header/structure from Kinder.
' -----------------------------------------------------------------------------
Private Function CreateKinderBlanksSheet(ByVal wsKinder As Worksheet) As Worksheet
    Dim wsBlanks As Worksheet
    Dim rngSource As Range
    Dim rngDest As Range
    Dim colIndex As Long
    Dim lastStructureCol As Long
    
    On Error GoTo ErrorHandler
    
    ' Determine last column index (T = 20)
    lastStructureCol = wsKinder.Range(LAST_STRUCTURE_COL & "1").Column
    
    ' Create new worksheet
    Set wsBlanks = ThisWorkbook.Worksheets.Add(After:=wsKinder)
    wsBlanks.name = SHEET_KINDER_BLANKS
    
    ' Copy header rows (1-4) and structure columns (A-T)
    Set rngSource = wsKinder.Range( _
        wsKinder.Cells(1, 1), _
        wsKinder.Cells(HEADER_LAST_ROW, lastStructureCol) _
    )
    Set rngDest = wsBlanks.Range( _
        wsBlanks.Cells(1, 1), _
        wsBlanks.Cells(HEADER_LAST_ROW, lastStructureCol) _
    )
    
    ' Copy values and formats (not formulas for data independence)
    rngSource.Copy
    rngDest.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    rngDest.PasteSpecial Paste:=xlPasteFormats
    rngDest.PasteSpecial Paste:=xlPasteColumnWidths
    Application.CutCopyMode = False
    
    ' Copy row heights for header rows
    Dim rowIdx As Long
    For rowIdx = 1 To HEADER_LAST_ROW
        wsBlanks.rows(rowIdx).RowHeight = wsKinder.rows(rowIdx).RowHeight
    Next rowIdx
    
    ' Ensure T2 is a plain value (not a formula reference)
    EnsureT2IsValue wsBlanks, wsKinder
    
    ' Clear any data that might have been copied beyond header
    ClearDataRows wsBlanks, FIRST_DATA_ROW, lastStructureCol
    
    Set CreateKinderBlanksSheet = wsBlanks
    Exit Function
    
ErrorHandler:
    ' Cleanup on error
    On Error Resume Next
    If Not wsBlanks Is Nothing Then
        Application.DisplayAlerts = False
        wsBlanks.Delete
        Application.DisplayAlerts = True
    End If
    Set CreateKinderBlanksSheet = Nothing
End Function

' -----------------------------------------------------------------------------
' EnsureT2IsValue
' Makes sure T2 in the blanks sheet is a plain value, not a formula.
' Copies the value from Kinder!T2 if needed.
' -----------------------------------------------------------------------------
Private Sub EnsureT2IsValue(ByVal wsBlanks As Worksheet, ByVal wsKinder As Worksheet)
    Dim cellT2 As Range
    Dim sourceValue As Variant
    
    Set cellT2 = wsBlanks.Range("T2")
    
    ' Get source value from Kinder
    sourceValue = wsKinder.Range("T2").value
    
    ' Clear any formula and set plain value
    cellT2.ClearContents
    
    If IsDate(sourceValue) Then
        cellT2.value = sourceValue
    ElseIf Not IsEmpty(sourceValue) Then
        cellT2.value = sourceValue
    End If
    
    ' Ensure date format
    cellT2.NumberFormat = "dd.mm.yyyy"
End Sub

' -----------------------------------------------------------------------------
' ClearDataRows
' Clears content from data rows (starting at firstDataRow) if any exist.
' -----------------------------------------------------------------------------
Private Sub ClearDataRows(ByVal ws As Worksheet, ByVal firstDataRow As Long, ByVal lastCol As Long)
    Dim lastRow As Long
    Dim rngData As Range
    
    ' Find last used row
    lastRow = ws.Cells(ws.rows.Count, 3).End(xlUp).row
    
    ' If data exists beyond header, clear it
    If lastRow >= firstDataRow Then
        Set rngData = ws.Range( _
            ws.Cells(firstDataRow, 1), _
            ws.Cells(lastRow, lastCol) _
        )
        rngData.ClearContents
    End If
End Sub
