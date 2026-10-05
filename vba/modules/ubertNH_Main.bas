Attribute VB_Name = "ubertNH_Main"
Option Explicit

' ====================================================================
' Module: ubertNH_Main
' Description: Main entry point for the transfer macro
' ====================================================================

' Main procedure - entry point
Public Sub ubertNH_TransferData()
    Dim srcWorkbook As Workbook
    Dim srcSheet As Worksheet
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet
    Dim records As Collection
    Dim stats As ubertNH_DataWriter.TransferStats
    Dim monthNumber As Integer
    Dim semester As Integer
    Dim originalCalculation As XlCalculation
    Dim originalScreenUpdating As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Save original Excel settings
    originalCalculation = Application.Calculation
    originalScreenUpdating = Application.ScreenUpdating
    
    ' Disable screen updating and automatic calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Get source workbook and sheet
    Set srcWorkbook = ThisWorkbook
    Set srcSheet = ActiveSheet
    
    ' Initialize logger
    On Error Resume Next
    ubertNH_Logger.InitializeLogger srcWorkbook.fullName
    If Err.Number <> 0 Then
        MsgBox "Fehler beim Initialisieren des Loggers: " & Err.Description, vbCritical
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' Clear filters on source sheet
    ubertNH_Utils.ClearFilters srcSheet
    
    ' Read transfer date and determine month/semester
    If Not ubertNH_DataReader.ReadTransferDate(srcWorkbook, monthNumber, semester) Then
        GoTo CleanupAndExit
    End If
    
    ubertNH_Logger.AddLogEntry "Monat: " & monthNumber & ", Semester: " & semester
    ubertNH_Logger.AddLogEntry ""
    
    ' Read source records
    ubertNH_Logger.AddLogEntry "Lese Quelldaten..."
    Set records = ubertNH_DataReader.ReadSourceRecords(srcSheet)
    ubertNH_Logger.AddLogEntry "Quelldaten gelesen: " & records.Count & " Datensaetze"
    
    ' Check if any data found
    If records.Count = 0 Then
        MsgBox ubertNH_Config.MSG_NO_DATA_FOUND, vbInformation
        GoTo CleanupAndExit
    End If
    
    ubertNH_Logger.AddLogEntry "Anzahl Quelldatensaetze: " & records.Count
    ubertNH_Logger.AddLogEntry ""
    
    ' Find open target workbook
    Set targetWorkbook = FindOpenWorkbook(ubertNH_Config.TARGET_FILE_NAME)
    If targetWorkbook Is Nothing Then
        MsgBox ubertNH_Config.MSG_TARGET_FILE_NOT_OPEN, vbExclamation
        GoTo CleanupAndExit
    End If
    
    ' Get target sheet
    On Error Resume Next
    Set targetSheet = targetWorkbook.Worksheets(ubertNH_Config.TARGET_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If targetSheet Is Nothing Then
        MsgBox ubertNH_Config.MSG_TARGET_SHEET_NOT_FOUND, vbExclamation
        GoTo CleanupAndExit
    End If
    
    ' Clear filters on target sheet
    ubertNH_Utils.ClearFilters targetSheet
    
    ' Process transfer
    stats = ubertNH_DataWriter.ProcessTransfer(srcSheet, targetSheet, records, monthNumber, semester)
    
    ' Save target workbook
    targetWorkbook.Save
    
    ' Write log to file
    ubertNH_Logger.WriteLogToFile
    
    ' Show completion message
    ShowCompletionMessage stats
    
CleanupAndExit:
    ' Restore Excel settings
    Application.Calculation = originalCalculation
    Application.ScreenUpdating = originalScreenUpdating
    Exit Sub
    
ErrorHandler:
    ' Restore Excel settings
    Application.Calculation = originalCalculation
    Application.ScreenUpdating = originalScreenUpdating
    
    ' Show detailed error message
    Dim errorMsg As String
    errorMsg = "FEHLER aufgetreten:" & vbCrLf & vbCrLf & _
               "Nummer: " & Err.Number & vbCrLf & _
               "Beschreibung: " & Err.Description & vbCrLf & _
               "Quelle: " & Err.Source & vbCrLf & vbCrLf & _
               "Bitte pruefen Sie die Details und versuchen Sie es erneut."
    
    MsgBox errorMsg, vbCritical, "Fehler"
    
    ' Try to write log
    On Error Resume Next
    ubertNH_Logger.AddLogEntry "FEHLER Nr. " & Err.Number & ": " & Err.Description
    ubertNH_Logger.AddLogEntry "Quelle: " & Err.Source
    ubertNH_Logger.WriteLogToFile
    On Error GoTo 0
End Sub

' Find open workbook by name
Private Function FindOpenWorkbook(ByVal workbookName As String) As Workbook
    Dim wb As Workbook
    
    ' Search through all open workbooks
    For Each wb In Application.Workbooks
        ' Compare names (case-insensitive)
        If StrComp(wb.name, workbookName, vbTextCompare) = 0 Then
            Set FindOpenWorkbook = wb
            Exit Function
        End If
    Next wb
    
    ' Not found
    Set FindOpenWorkbook = Nothing
End Function

' Show completion message with statistics
Private Sub ShowCompletionMessage(ByRef stats As ubertNH_DataWriter.TransferStats)
    Dim msg As String
    
    msg = ubertNH_Config.MSG_PROCESS_COMPLETE
    msg = ubertNH_Utils.FormatMessage(msg, _
                                      stats.successCount, _
                                      stats.notFoundCount, _
                                      stats.overwriteCount, _
                                      ubertNH_Logger.GetLogFilePath())
    
    MsgBox msg, vbInformation
End Sub
