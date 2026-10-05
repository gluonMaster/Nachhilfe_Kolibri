Attribute VB_Name = "NHblank_RecordView"
Option Explicit

' =============================================================================
' NHblank_RecordView
' -----------------------------------------------------------------------------
' Provides functions for checking record activity and formatting records
' for user display (e.g., duplicate resolution dialogs).
'
' Active record rule: Column C is not empty AND Font.ColorIndex <> 15 (not gray)
' =============================================================================

Private Const GRAY_FONT_COLOR_INDEX As Long = 15

' -----------------------------------------------------------------------------
' NHblank_IsActiveRecord
' -----------------------------------------------------------------------------
' Checks if a record is considered "active" based on column C.
'
' Parameters:
'   ws      - Worksheet containing the record
'   rowNum  - Row number to check
'
' Returns:
'   True if column C is not empty and font color is not gray (ColorIndex 15)
' -----------------------------------------------------------------------------
Public Function NHblank_IsActiveRecord(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    Dim cellC As Range
    Set cellC = ws.Cells(rowNum, "C")
    
    ' Check if cell is empty
    If IsEmpty(cellC.value) Then
        NHblank_IsActiveRecord = False
        Exit Function
    End If
    
    If Len(Trim(CStr(cellC.value))) = 0 Then
        NHblank_IsActiveRecord = False
        Exit Function
    End If
    
    ' Check if font color is gray (inactive)
    If cellC.Font.ColorIndex = GRAY_FONT_COLOR_INDEX Then
        NHblank_IsActiveRecord = False
        Exit Function
    End If
    
    NHblank_IsActiveRecord = True
End Function

' -----------------------------------------------------------------------------
' NHblank_FormatRecordForChoice
' -----------------------------------------------------------------------------
' Formats a record as a readable string for user display.
' Includes key fields (B,C,D,E,F,I,K,L,O), date range (G/H), and S/T fields.
'
' Parameters:
'   ws      - Worksheet containing the record
'   rowNum  - Row number to format
'
' Returns:
'   Formatted string representation of the record
' -----------------------------------------------------------------------------
Public Function NHblank_FormatRecordForChoice(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    Dim parts As String
    
    ' Key fields: B, C, D, E, F, I, K, L, O
    parts = "Zeile " & rowNum & ": "
    parts = parts & "B=" & SafeValue(ws.Cells(rowNum, "B").value) & ", "
    parts = parts & "C=" & SafeValue(ws.Cells(rowNum, "C").value) & ", "
    parts = parts & "D=" & SafeValue(ws.Cells(rowNum, "D").value) & ", "
    parts = parts & "E=" & SafeValue(ws.Cells(rowNum, "E").value) & ", "
    parts = parts & "F=" & SafeValue(ws.Cells(rowNum, "F").value) & ", "
    parts = parts & "I=" & SafeValue(ws.Cells(rowNum, "I").value) & ", "
    parts = parts & "K=" & SafeValue(ws.Cells(rowNum, "K").value) & ", "
    parts = parts & "L=" & SafeValue(ws.Cells(rowNum, "L").value) & ", "
    parts = parts & "O=" & SafeValue(ws.Cells(rowNum, "O").value)
    
    ' Date range: G (von) and H (bis)
    parts = parts & " | Daten: "
    parts = parts & FormatDateValue(ws.Cells(rowNum, "G").value)
    parts = parts & " - "
    parts = parts & FormatDateValue(ws.Cells(rowNum, "H").value)
    
    ' Additional fields: S and T
    parts = parts & " | S=" & SafeValue(ws.Cells(rowNum, "S").value)
    parts = parts & ", T=" & SafeValue(ws.Cells(rowNum, "T").value)
    
    NHblank_FormatRecordForChoice = parts
End Function

' -----------------------------------------------------------------------------
' SafeValue
' -----------------------------------------------------------------------------
' Converts a cell value to a safe string representation.
' Handles Empty, Error, and Null values.
'
' Parameters:
'   cellValue - Value from worksheet cell (Variant)
'
' Returns:
'   String representation, "(empty)" for empty values, "(error)" for errors
' -----------------------------------------------------------------------------
Private Function SafeValue(ByVal cellValue As Variant) As String
    If IsEmpty(cellValue) Then
        SafeValue = "(leer)"
        Exit Function
    End If
    
    If IsError(cellValue) Then
        SafeValue = "(fehler)"
        Exit Function
    End If
    
    If IsNull(cellValue) Then
        SafeValue = "(null)"
        Exit Function
    End If
    
    Dim strVal As String
    strVal = Trim(CStr(cellValue))
    
    If Len(strVal) = 0 Then
        SafeValue = "(leer)"
    Else
        SafeValue = strVal
    End If
End Function

' -----------------------------------------------------------------------------
' FormatDateValue
' -----------------------------------------------------------------------------
' Formats a date value as dd.mm.yyyy if it's a valid date,
' otherwise returns the value as-is.
'
' Parameters:
'   cellValue - Value from worksheet cell (Variant)
'
' Returns:
'   Formatted date string or original value as string
' -----------------------------------------------------------------------------
Private Function FormatDateValue(ByVal cellValue As Variant) As String
    If IsEmpty(cellValue) Then
        FormatDateValue = "(leer)"
        Exit Function
    End If
    
    If IsError(cellValue) Then
        FormatDateValue = "(fehler)"
        Exit Function
    End If
    
    If IsNull(cellValue) Then
        FormatDateValue = "(null)"
        Exit Function
    End If
    
    ' Check if it's a date
    If IsDate(cellValue) Then
        FormatDateValue = Format(CDate(cellValue), "dd.mm.yyyy")
    Else
        ' Return as-is if not a date
        Dim strVal As String
        strVal = Trim(CStr(cellValue))
        If Len(strVal) = 0 Then
            FormatDateValue = "(leer)"
        Else
            FormatDateValue = strVal
        End If
    End If
End Function
