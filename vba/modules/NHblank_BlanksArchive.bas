Attribute VB_Name = "NHblank_BlanksArchive"
Option Explicit

' =============================================================================
' NHblank_BlanksArchive
' Deactivates records on Kinder_Blanks based on reference date from T2.
' Does NOT move/delete records, only changes font color (gray = inactive).
' =============================================================================

Private Const FIRST_DATA_ROW As Long = 5
Private Const COL_C As Long = 3
Private Const COL_G As Long = 7
Private Const COL_H As Long = 8
Private Const COL_T As Long = 20
Private Const LAST_FORMAT_COL As String = "T"

' Font color constants
Private Const FONT_COLOR_ACTIVE As Long = -4105   ' xlAutomatic
Private Const FONT_COLOR_INACTIVE As Long = 15    ' Gray

' -----------------------------------------------------------------------------
' NHblank_DeactivateBlanksByT2
' Main entry point: deactivates records on Kinder_Blanks based on T2 date.
' Records outside the reference month are marked as inactive (gray font).
' -----------------------------------------------------------------------------
Public Sub NHblank_DeactivateBlanksByT2()
    Dim appState As NHblank_AppState
    Dim wsBlanks As Worksheet
    Dim referenceDate As Variant
    Dim referenceMonth As Integer
    Dim referenceYear As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim dateFrom As Variant
    Dim dateTo As Variant
    Dim activeCount As Long
    Dim inactiveCount As Long
    Dim rowRange As Range
    
    ' Initialize app state for safe cleanup
    Set appState = New NHblank_AppState
    appState.Capture
    appState.OptimizeForRun
    
    On Error GoTo ErrorHandler
    
    ' Get Kinder_Blanks worksheet (must already exist)
    Set wsBlanks = NHblank_Sheets.NHblank_GetKinderBlanksSheet()
    
    If wsBlanks Is Nothing Then
        MsgBox "FEHLER: Worksheet 'Kinder_Blanks' nicht gefunden!" & vbCrLf & _
               "Bitte zuerst die Synchronisation ausfuehren.", _
               vbExclamation, "Kinder_Blanks fehlt"
        GoTo CleanExit
    End If
    
    ' Apply format guards before processing
    NHblank_FormatGuard.NHblank_ApplyAllFormats wsBlanks, FIRST_DATA_ROW
    
    ' Get reference date from Kinder_Blanks!T2
    On Error Resume Next
    referenceDate = wsBlanks.Cells(2, COL_T).value
    On Error GoTo ErrorHandler
    
    If Not IsDate(referenceDate) Then
        MsgBox "FEHLER: Zelle T2 auf 'Kinder_Blanks' enthaelt kein gueltiges Datum!" & vbCrLf & _
               "Bitte ein Referenzdatum in T2 eintragen.", _
               vbExclamation, "Referenzdatum fehlt"
        GoTo CleanExit
    End If
    
    ' Extract month and year from reference date
    referenceMonth = Month(CDate(referenceDate))
    referenceYear = Year(CDate(referenceDate))
    
    ' Determine last data row
    lastRow = wsBlanks.Cells(wsBlanks.rows.Count, COL_C).End(xlUp).row
    
    If lastRow < FIRST_DATA_ROW Then
        MsgBox "Keine Datensaetze auf 'Kinder_Blanks' gefunden.", _
               vbInformation, "Keine Daten"
        GoTo CleanExit
    End If
    
    ' Initialize counters
    activeCount = 0
    inactiveCount = 0
    
    ' Process each row
    For i = FIRST_DATA_ROW To lastRow
        ' Skip empty rows (column C is primary identifier)
        If Trim(wsBlanks.Cells(i, COL_C).value) = "" Then
            GoTo NextRow
        End If
        
        ' Read date range from columns G and H
        dateFrom = wsBlanks.Cells(i, COL_G).value
        dateTo = wsBlanks.Cells(i, COL_H).value
        
        ' Get row range for formatting (A:T)
        Set rowRange = wsBlanks.Range("A" & i & ":" & LAST_FORMAT_COL & i)
        
        ' Check if record is active based on date range
        If IsDateRangeActiveForMonth(dateFrom, dateTo, referenceMonth, referenceYear) Then
            ' Mark as active (automatic/black font)
            rowRange.Font.ColorIndex = FONT_COLOR_ACTIVE
            activeCount = activeCount + 1
        Else
            ' Mark as inactive (gray font)
            rowRange.Font.ColorIndex = FONT_COLOR_INACTIVE
            inactiveCount = inactiveCount + 1
        End If
        
NextRow:
    Next i
    
    ' Sort Kinder_Blanks by column C ascending
    SortByColumnC wsBlanks

    ' Show summary message (German, no umlauts)
    MsgBox "Deaktivierung abgeschlossen." & vbCrLf & vbCrLf & _
           "Referenzdatum: " & Format(referenceDate, "dd.mm.yyyy") & vbCrLf & _
           "Aktive Datensaetze: " & activeCount & vbCrLf & _
           "Inaktive Datensaetze (grau): " & inactiveCount, _
           vbInformation, "Kinder_Blanks Deaktivierung"
    
CleanExit:
    appState.Restore
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler bei der Deaktivierung: " & Err.Description, _
           vbCritical, "Fehler"
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------
' IsDateRangeActiveForMonth
' Determines if a date range is considered "active" for the given month/year.
' Logic mirrors Archivierung.IsDateRangeActive but is self-contained.
'
' A record is ACTIVE if:
'   1) The date range (G to H) spans/covers the reference month, OR
'   2) The start date (G) is in a future month relative to reference date
'
' Returns False (inactive) otherwise.
' -----------------------------------------------------------------------------
Private Function IsDateRangeActiveForMonth( _
    ByVal startDate As Variant, _
    ByVal endDate As Variant, _
    ByVal refMonth As Integer, _
    ByVal refYear As Integer _
) As Boolean
    
    Dim startDt As Date
    Dim endDt As Date
    Dim refMonthStart As Date
    
    IsDateRangeActiveForMonth = False
    
    ' Both dates must be valid
    If Not IsDate(startDate) Or Not IsDate(endDate) Then
        Exit Function
    End If
    
    startDt = CDate(startDate)
    endDt = CDate(endDate)
    refMonthStart = DateSerial(refYear, refMonth, 1)
    
    ' Case 1: Date range spans the reference month
    ' - Start date is in reference month, OR
    ' - End date is in reference month, OR
    ' - Range straddles the reference month (start before, end after)
    If (Year(startDt) = refYear And Month(startDt) = refMonth) Or _
       (Year(endDt) = refYear And Month(endDt) = refMonth) Or _
       (startDt <= refMonthStart And endDt >= refMonthStart) Then
        IsDateRangeActiveForMonth = True
        Exit Function
    End If
    
    ' Case 2: Start date is in a future month relative to reference
    If (Year(startDt) > refYear) Or _
       (Year(startDt) = refYear And Month(startDt) > refMonth) Then
        IsDateRangeActiveForMonth = True
        Exit Function
    End If
    
    ' Otherwise inactive (date range is in the past)
    IsDateRangeActiveForMonth = False
End Function

' -----------------------------------------------------------------------------
' SortByColumnC
' -----------------------------------------------------------------------------
' Sorts data rows (from FIRST_DATA_ROW) on Kinder_Blanks by column C ascending.
' -----------------------------------------------------------------------------
Private Sub SortByColumnC(ByVal ws As Worksheet)
    Dim lastRow As Long

    lastRow = ws.Cells(ws.rows.Count, COL_C).End(xlUp).row
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
