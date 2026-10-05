Attribute VB_Name = "NHblank_DuplicateResolver"
Option Explicit

' =============================================================================
' NHblank_DuplicateResolver
' -----------------------------------------------------------------------------
' Provides UI for resolving duplicate records when multiple rows share
' the same key during synchronization operations.
'
' Uses MsgBox + InputBox approach (no UserForm).
' =============================================================================

' -----------------------------------------------------------------------------
' NHblank_ResolveDuplicateRow
' -----------------------------------------------------------------------------
' Prompts the user to select one record from a collection of duplicate rows.
'
' Parameters:
'   ws   - Worksheet containing the records
'   key  - The duplicate key value (for display purposes)
'   rows - Collection of row numbers (Long) that share the same key
'
' Returns:
'   Selected row number (Long), or 0 if user cancels/skips
' -----------------------------------------------------------------------------
Public Function NHblank_ResolveDuplicateRow(ByVal ws As Worksheet, _
                                            ByVal key As String, _
                                            ByVal rows As Collection) As Long
    Dim msg As String
    Dim i As Long
    Dim rowNum As Long
    Dim userInput As String
    Dim selectedIndex As Long
    
    ' Validate input
    If rows Is Nothing Then
        NHblank_ResolveDuplicateRow = 0
        Exit Function
    End If
    
    If rows.Count = 0 Then
        NHblank_ResolveDuplicateRow = 0
        Exit Function
    End If
    
    ' If only one row, return it directly (no need for user choice)
    If rows.Count = 1 Then
        NHblank_ResolveDuplicateRow = CLng(rows(1))
        Exit Function
    End If
    
    ' Build the message with all duplicate records
    msg = "Mehrere Datensaetze mit gleichem Schluessel gefunden:" & vbCrLf & vbCrLf
    msg = msg & "Schluessel: " & key & vbCrLf & vbCrLf
    
    For i = 1 To rows.Count
        rowNum = CLng(rows(i))
        msg = msg & i & " - " & NHblank_FormatRecordForChoice(ws, rowNum) & vbCrLf
    Next i
    
    msg = msg & vbCrLf & "Bitte waehlen Sie eine Nummer (1-" & rows.Count & ")."
    msg = msg & vbCrLf & "Druecken Sie Abbrechen, um diesen Datensatz zu ueberspringen."
    
    ' Show the list of duplicates
    MsgBox msg, vbInformation, "Duplikate gefunden"
    
    ' Ask user to select a number
    userInput = InputBox("Geben Sie die Nummer des gewuenschten Datensatzes ein (1-" & rows.Count & "):" & vbCrLf & vbCrLf & _
                         "Leer lassen oder Abbrechen = ueberspringen", _
                         "Datensatz auswaehlen", "")
    
    ' Handle cancel or empty input
    If Len(Trim(userInput)) = 0 Then
        NHblank_ResolveDuplicateRow = 0
        Exit Function
    End If
    
    ' Validate numeric input
    If Not IsNumeric(userInput) Then
        MsgBox "Ungueltige Eingabe. Datensatz wird uebersprungen.", vbExclamation, "Fehler"
        NHblank_ResolveDuplicateRow = 0
        Exit Function
    End If
    
    selectedIndex = CLng(userInput)
    
    ' Validate range
    If selectedIndex < 1 Or selectedIndex > rows.Count Then
        MsgBox "Nummer ausserhalb des gueltigen Bereichs (1-" & rows.Count & ")." & vbCrLf & _
               "Datensatz wird uebersprungen.", vbExclamation, "Fehler"
        NHblank_ResolveDuplicateRow = 0
        Exit Function
    End If
    
    ' Return the selected row number
    NHblank_ResolveDuplicateRow = CLng(rows(selectedIndex))
End Function
