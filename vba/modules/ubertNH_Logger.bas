Attribute VB_Name = "ubertNH_Logger"
Option Explicit

' ====================================================================
' Module: ubertNH_Logger
' Description: Logging functionality for transfer operations
' ====================================================================

Private m_logFilePath As String
Private m_logEntries As Collection

' Initialize logger
Public Sub InitializeLogger(ByVal sourceFilePath As String)
    Dim folderPath As String
    Dim fileName As String
    
    ' Get folder path from source file
    folderPath = Left(sourceFilePath, InStrRev(sourceFilePath, "\"))
    
    ' Create log file name with timestamp
    fileName = ubertNH_Config.LOG_FILE_PREFIX & _
               ubertNH_Utils.GetTimestamp() & _
               ubertNH_Config.LOG_FILE_EXTENSION
    
    m_logFilePath = folderPath & fileName
    
    ' Initialize collection
    Set m_logEntries = New Collection
    
    ' Add header
    AddLogEntry "==================================================================="
    AddLogEntry "Transfer Log - " & Format(Now, "DD.MM.YYYY HH:MM:SS")
    AddLogEntry "==================================================================="
    AddLogEntry ""
End Sub

' Add entry to log
Public Sub AddLogEntry(ByVal entry As String)
    If m_logEntries Is Nothing Then
        Set m_logEntries = New Collection
    End If
    
    m_logEntries.Add entry
End Sub

' Log record not found with reason
Public Sub LogRecordNotFound(ByVal sourceRow As Long, _
                             ByVal ID As String, _
                             ByVal fullName As String, _
                             ByVal subject As String, _
                             ByVal reason As String)
    AddLogEntry "--- Datensatz nicht gefunden (Zeile " & sourceRow & ") ---"
    AddLogEntry "  ID: " & ID
    AddLogEntry "  Name: " & fullName
    AddLogEntry "  Fach: " & subject
    AddLogEntry "  Grund: " & reason
    AddLogEntry ""
End Sub

' Log value overwrite
Public Sub LogValueOverwrite(ByVal targetSheet As String, _
                            ByVal targetRow As Long, _
                            ByVal targetCol As String, _
                            ByVal ID As String, _
                            ByVal fullName As String, _
                            ByVal subject As String, _
                            ByVal oldValue As Variant, _
                            ByVal newValue As Variant)
    AddLogEntry "--- Wert ueberschrieben ---"
    AddLogEntry "  Zelle: " & targetSheet & "!" & targetCol & targetRow
    AddLogEntry "  ID: " & ID
    AddLogEntry "  Name: " & fullName
    AddLogEntry "  Fach: " & subject
    AddLogEntry "  Alter Wert: " & CStr(oldValue)
    AddLogEntry "  Neuer Wert: " & CStr(newValue)
    AddLogEntry ""
End Sub

' Write log to file
Public Sub WriteLogToFile()
    Dim fileNum As Integer
    Dim entry As Variant
    
    If m_logEntries Is Nothing Then Exit Sub
    If m_logEntries.Count = 0 Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    fileNum = FreeFile
    Open m_logFilePath For Output As #fileNum
    
    For Each entry In m_logEntries
        Print #fileNum, entry
    Next entry
    
    Close #fileNum
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    Close #fileNum
    MsgBox "Fehler beim Schreiben der Log-Datei: " & Err.Description, vbExclamation
End Sub

' Get log file path
Public Function GetLogFilePath() As String
    GetLogFilePath = m_logFilePath
End Function

' Check if log has entries (beyond header)
Public Function HasLogEntries() As Boolean
    If m_logEntries Is Nothing Then
        HasLogEntries = False
    Else
        ' Header has 4 entries, so check if more than that
        HasLogEntries = (m_logEntries.Count > 4)
    End If
End Function
