Attribute VB_Name = "ubertNH_Config"
Option Explicit

' ====================================================================
' Module: ubertNH_Config
' Description: Configuration constants and settings for the transfer macro
' ====================================================================

' Source file settings
Public Const SRC_FIRST_DATA_ROW As Long = 11
Public Const SRC_COL_ID As String = "A"
Public Const SRC_COL_LASTNAME As String = "B"
Public Const SRC_COL_FIRSTNAME As String = "C"
Public Const SRC_COL_SUBJECT As String = "D"
Public Const SRC_COL_VALUE As String = "AQ"  ' Column with values to transfer

' Target file settings
Public Const TARGET_FILE_NAME As String = "KindElternDaten_26_Admin.xlsm"
Public Const TGT_FIRST_DATA_ROW As Long = 3
Public Const TGT_COL_ID As String = "A"
Public Const TGT_COL_ID_CHECK As String = "B"
Public Const TGT_COL_FULLNAME As String = "D"
Public Const TGT_COL_SUBJECT_SEMESTER1 As String = "J"
Public Const TGT_COL_SUBJECT_SEMESTER2 As String = "O"
Public Const TGT_COL_MONTH_START As String = "U"  ' January
Public Const TGT_COL_MONTH_END As String = "AF"   ' December

' Date settings
Public Const DATE_SHEET_NAME As String = "Kinder"
Public Const DATE_CELL As String = "T2"

' Semester definitions (months)
Public Const SEMESTER1_MONTHS As String = "1,2,3,4,5,6"  ' January - June
Public Const SEMESTER2_MONTHS As String = "8,9,10,11,12" ' August - December

' Target sheet name
Public Const TARGET_SHEET_NAME As String = "Kartei"

' Highlight color for not found records (pale pink)
Public Const HIGHLIGHT_COLOR_RGB As Long = 13421772  ' RGB(255, 200, 200) converted to Long

' Log file settings
Public Const LOG_FILE_PREFIX As String = "ubertNH_Transfer_Log_"
Public Const LOG_FILE_EXTENSION As String = ".txt"

' Messages in German (without umlauts)
Public Const MSG_TARGET_FILE_NOT_OPEN As String = "Die Zieldatei 'KindElternDaten_26_Admin.xlsm' ist nicht geoeffnet." & vbCrLf & _
                                                   "Bitte oeffnen Sie die Datei und fuehren Sie das Makro erneut aus."
Public Const MSG_TARGET_SHEET_NOT_FOUND As String = "Zielblatt 'Kartei' wurde nicht gefunden."
Public Const MSG_DATE_SHEET_NOT_FOUND As String = "Blatt 'Kinder' wurde nicht gefunden."
Public Const MSG_DATE_CELL_EMPTY As String = "Datumszelle ist leer."
Public Const MSG_INVALID_DATE As String = "Ungueltiges Datum in Zelle T2."
Public Const MSG_NO_DATA_FOUND As String = "Keine Daten zum Uebertragen gefunden."
Public Const MSG_PROCESS_COMPLETE As String = "Uebertragung abgeschlossen!" & vbCrLf & vbCrLf & _
                                               "Erfolgreich uebertragen: {SUCCESS}" & vbCrLf & _
                                               "Nicht gefunden: {NOTFOUND}" & vbCrLf & _
                                               "Ueberschreibungen: {OVERWRITES}" & vbCrLf & vbCrLf & _
                                               "Log-Datei: {LOGPATH}"
Public Const MSG_ERROR_OCCURRED As String = "Fehler aufgetreten: {ERROR}"
