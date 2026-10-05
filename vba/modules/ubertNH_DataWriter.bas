Attribute VB_Name = "ubertNH_DataWriter"
Option Explicit

' ====================================================================
' Module: ubertNH_DataWriter
' Description: Write data to target worksheet
' ====================================================================

' Structure to hold transfer statistics
Public Type TransferStats
    successCount As Long
    notFoundCount As Long
    overwriteCount As Long
End Type

' Write value to target cell and track overwrites
Public Sub WriteValueToTarget(ByVal targetSheet As Worksheet, _
                             ByVal targetRow As Long, _
                             ByVal monthNumber As Integer, _
                             ByRef srcRecord As SourceRecord, _
                             ByRef stats As TransferStats)
    Dim targetCol As String
    Dim targetCell As Range
    Dim oldValue As Variant
    
    ' Get column letter for the month
    targetCol = ubertNH_Utils.GetMonthColumn(monthNumber)
    
    ' Get target cell
    Set targetCell = targetSheet.Range(targetCol & targetRow)
    
    ' Read old value
    oldValue = targetCell.value
    
    ' Check if overwriting non-zero value with DIFFERENT value
    If Not ubertNH_Utils.IsCellEmptyOrZero(oldValue) Then
        ' Check if values are different
        If oldValue <> srcRecord.value Then
            ' Log overwrite only if values differ
            ubertNH_Logger.LogValueOverwrite _
                targetSheet.name, _
                targetRow, _
                targetCol, _
                srcRecord.ID, _
                srcRecord.lastName & " " & srcRecord.firstName, _
                srcRecord.subject, _
                oldValue, _
                srcRecord.value
            
            stats.overwriteCount = stats.overwriteCount + 1
        End If
    End If
    
    ' Write new value
    targetCell.value = srcRecord.value
    
    ' Increment success counter
    stats.successCount = stats.successCount + 1
End Sub

' Highlight source row as not found
Public Sub HighlightNotFoundRow(ByVal srcSheet As Worksheet, _
                               ByRef srcRecord As SourceRecord)
    Dim HighlightRange As Range
    
    ' Highlight cells A:C in the source row
    Set HighlightRange = srcSheet.Range( _
        ubertNH_Config.SRC_COL_ID & srcRecord.RowNumber & ":" & _
        ubertNH_Config.SRC_COL_FIRSTNAME & srcRecord.RowNumber)
    
    ubertNH_Utils.HighlightRange HighlightRange
End Sub

' Process all records transfer
Public Function ProcessTransfer(ByVal srcSheet As Worksheet, _
                               ByVal targetSheet As Worksheet, _
                               ByVal records As Collection, _
                               ByVal monthNumber As Integer, _
                               ByVal semester As Integer) As TransferStats
    Dim stats As TransferStats
    Dim record As SourceRecord
    Dim FindResult As FindResult
    Dim i As Long
    
    ' Initialize stats
    stats.successCount = 0
    stats.notFoundCount = 0
    stats.overwriteCount = 0
    
    ' Process each record
    For i = 1 To records.Count
        Set record = records(i)
        
        ' Find matching target record
        Set FindResult = ubertNH_Finder.FindTargetRecord(targetSheet, record, semester)
        
        If FindResult.found Then
            ' Write value to target
            WriteValueToTarget targetSheet, FindResult.targetRow, monthNumber, record, stats
        Else
            ' Log not found
            ubertNH_Logger.LogRecordNotFound _
                record.RowNumber, _
                record.ID, _
                record.lastName & " " & record.firstName, _
                record.subject, _
                FindResult.notFoundReason
            
            ' Highlight source row
            HighlightNotFoundRow srcSheet, record
            
            stats.notFoundCount = stats.notFoundCount + 1
        End If
    Next i
    
    ProcessTransfer = stats
End Function
