Attribute VB_Name = "NHblank_FormatGuard"
Option Explicit

' =============================================================================
' NHblank_FormatGuard
' Utilities to protect BG-Nummer (column O) as text and ensure date formats.
' Reusable across project modules; does not depend on specific sheet names.
' =============================================================================

' Column constants (1-based)
Private Const COL_B As Long = 2
Private Const COL_C As Long = 3
Private Const COL_G As Long = 7
Private Const COL_H As Long = 8
Private Const COL_O As Long = 15
Private Const COL_T As Long = 20

' -----------------------------------------------------------------------------
' NHblank_EnsureBgNummerTextFormat
' Sets column O to text format (@) for the entire column.
' Prevents Excel from auto-converting BG-Nummer to dates, exponential, etc.
' Always applies format regardless of whether data exists.
' -----------------------------------------------------------------------------
Public Sub NHblank_EnsureBgNummerTextFormat(ByVal ws As Worksheet, ByVal firstDataRow As Long)
    If ws Is Nothing Then Exit Sub
    
    ' Set text format for entire column O to prevent auto-conversion
    ws.Columns(COL_O).NumberFormat = "@"
End Sub

' -----------------------------------------------------------------------------
' NHblank_WriteBgNummer
' Safely writes a BG-Nummer value to a cell, preserving it as text.
' Prevents loss of leading zeros, auto-date conversion, exponential format.
' Uses NHblank_BgNummerToString for stable string conversion.
' -----------------------------------------------------------------------------
Public Sub NHblank_WriteBgNummer(ByVal targetCell As Range, ByVal bgValue As Variant)
    If targetCell Is Nothing Then Exit Sub
    
    ' Set text format first
    targetCell.NumberFormat = "@"
    
    ' Write using safe string conversion
    targetCell.Value2 = NHblank_BgNummerToString(bgValue)
End Sub

' -----------------------------------------------------------------------------
' NHblank_BgNummerToString
' Converts a BG-Nummer value to a stable string representation.
' Handles Empty/Null/Error gracefully, avoids exponential notation for numbers.
'
' Parameters:
'   bgValue - Variant value (from cell or variable)
'
' Returns:
'   String representation suitable for text storage
' -----------------------------------------------------------------------------
Public Function NHblank_BgNummerToString(ByVal bgValue As Variant) As String
    ' Handle Empty, Null, Error
    If IsEmpty(bgValue) Then
        NHblank_BgNummerToString = ""
        Exit Function
    End If
    
    If IsNull(bgValue) Then
        NHblank_BgNummerToString = ""
        Exit Function
    End If
    
    If IsError(bgValue) Then
        NHblank_BgNummerToString = ""
        Exit Function
    End If
    
    ' Check if numeric and integer-like (to avoid exponential notation)
    If IsNumeric(bgValue) Then
        Dim dblVal As Double
        dblVal = CDbl(bgValue)
        ' Check if it's an integer value (no fractional part)
        If dblVal = Int(dblVal) Then
            ' Use Format$ with "0" to avoid exponential notation
            NHblank_BgNummerToString = Trim$(Format$(dblVal, "0"))
            Exit Function
        End If
    End If
    
    ' Default: convert to string and trim
    NHblank_BgNummerToString = Trim$(CStr(bgValue))
End Function

' -----------------------------------------------------------------------------
' NHblank_EnsureDateDisplayFormat
' Sets date display format (dd.mm.yyyy) for columns G, H and cell T2.
' Dates remain stored as Date/Variant-Date, only display format changes.
' Always applies format regardless of whether data exists.
' -----------------------------------------------------------------------------
Public Sub NHblank_EnsureDateDisplayFormat(ByVal ws As Worksheet, ByVal firstDataRow As Long)
    If ws Is Nothing Then Exit Sub
    
    ' Format entire columns G and H for date display
    ws.Columns(COL_G).NumberFormat = "dd.mm.yyyy"
    ws.Columns(COL_H).NumberFormat = "dd.mm.yyyy"
    
    ' Format reference date cell T2 (always row 2)
    ws.Cells(2, COL_T).NumberFormat = "dd.mm.yyyy"
End Sub

' -----------------------------------------------------------------------------
' NHblank_ApplyAllFormats
' Convenience method: applies both BG-Nummer and date formats in one call.
' -----------------------------------------------------------------------------
Public Sub NHblank_ApplyAllFormats(ByVal ws As Worksheet, ByVal firstDataRow As Long)
    NHblank_EnsureBgNummerTextFormat ws, firstDataRow
    NHblank_EnsureDateDisplayFormat ws, firstDataRow
End Sub

