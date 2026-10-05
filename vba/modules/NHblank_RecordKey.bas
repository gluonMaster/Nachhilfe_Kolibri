Attribute VB_Name = "NHblank_RecordKey"
Option Explicit

' =============================================================================
' NHblank_RecordKey
' -----------------------------------------------------------------------------
' Builds a composite key from record columns for comparison between
' Kinder and Kinder_Blanks sheets.
'
' Key columns: B, C, D, E, F, I, K, L, O
' Excludes: G/H (dates), S/T, T2 (reference date)
' =============================================================================

Private Const KEY_DELIMITER As String = "|"

' -----------------------------------------------------------------------------
' NHblank_BuildRecordKey
' -----------------------------------------------------------------------------
' Builds a normalized composite key from the specified row.
'
' Parameters:
'   ws      - Worksheet containing the record
'   rowNum  - Row number to read
'
' Returns:
'   Lowercase string key with pipe-delimited column values.
'   Returns stable key even for partially empty rows.
' -----------------------------------------------------------------------------
Public Function NHblank_BuildRecordKey(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    Dim keyParts(1 To 9) As String
    
    ' Column B - typically identifier or number
    keyParts(1) = NormalizeCell(ws.Cells(rowNum, "B").value)
    
    ' Column C - Nachname (last name)
    keyParts(2) = NormalizeCell(ws.Cells(rowNum, "C").value)
    
    ' Column D - Vorname (first name)
    keyParts(3) = NormalizeCell(ws.Cells(rowNum, "D").value)
    
    ' Column E - Fach (subject/discipline)
    keyParts(4) = NormalizeCell(ws.Cells(rowNum, "E").value)
    
    ' Column F - additional identifier
    keyParts(5) = NormalizeCell(ws.Cells(rowNum, "F").value)
    
    ' Column I - Lehrer (teacher)
    keyParts(6) = NormalizeCell(ws.Cells(rowNum, "I").value)
    
    ' Column K - additional field
    keyParts(7) = NormalizeCell(ws.Cells(rowNum, "K").value)
    
    ' Column L - additional field
    keyParts(8) = NormalizeCell(ws.Cells(rowNum, "L").value)
    
    ' Column O - BG-Nummer (preserve as text, no auto-conversion)
    keyParts(9) = NormalizeCell(ws.Cells(rowNum, "O").value)
    
    ' Join all parts with delimiter and convert to lowercase for case-insensitive comparison
    NHblank_BuildRecordKey = LCase(Join(keyParts, KEY_DELIMITER))
End Function

' -----------------------------------------------------------------------------
' NormalizeCell
' -----------------------------------------------------------------------------
' Converts a cell value to a stable string representation.
' Handles Empty, Error, and normal values.
'
' Parameters:
'   cellValue - Value from worksheet cell (Variant)
'
' Returns:
'   Trimmed string representation, empty string for Empty/Error values.
' -----------------------------------------------------------------------------
Private Function NormalizeCell(ByVal cellValue As Variant) As String
    ' Handle Empty
    If IsEmpty(cellValue) Then
        NormalizeCell = ""
        Exit Function
    End If
    
    ' Handle Error values (e.g., #N/A, #REF!, #VALUE!)
    If IsError(cellValue) Then
        NormalizeCell = ""
        Exit Function
    End If
    
    ' Handle Null (rare in Excel, but possible)
    If IsNull(cellValue) Then
        NormalizeCell = ""
        Exit Function
    End If
    
    ' Convert to string and trim whitespace
    NormalizeCell = Trim(CStr(cellValue))
End Function
