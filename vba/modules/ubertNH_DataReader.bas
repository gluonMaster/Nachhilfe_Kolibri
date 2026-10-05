Attribute VB_Name = "ubertNH_DataReader"
Option Explicit

' ====================================================================
' Module: ubertNH_DataReader
' Description: Read data from source worksheet
' ====================================================================

' Read all source records from active sheet
Public Function ReadSourceRecords(ByVal srcSheet As Worksheet) As Collection
    Dim records As Collection
    Dim lastRow As Long
    Dim currentRow As Long
    Dim record As SourceRecord
    Dim idValue As String
    Dim aqValue As Variant
    
    Set records = New Collection
    
    ' Get last row based on column A
    lastRow = ubertNH_Utils.GetLastRow(srcSheet, ubertNH_Config.SRC_COL_ID)
    
    ' Loop through rows starting from first data row
    For currentRow = ubertNH_Config.SRC_FIRST_DATA_ROW To lastRow
        ' Check if column A is not empty
        idValue = Trim(CStr(srcSheet.Range(ubertNH_Config.SRC_COL_ID & currentRow).value))
        
        If idValue <> "" Then
            ' Check if AQ value is not empty
            aqValue = srcSheet.Range(ubertNH_Config.SRC_COL_VALUE & currentRow).value

            If Not ubertNH_Utils.IsCellEmptyOrZero(aqValue) Then
                ' Create new record object
                Set record = New SourceRecord
                
                ' Read record
                With record
                    .RowNumber = currentRow
                    .ID = idValue
                    .lastName = Trim(CStr(srcSheet.Range(ubertNH_Config.SRC_COL_LASTNAME & currentRow).value))
                    .firstName = Trim(CStr(srcSheet.Range(ubertNH_Config.SRC_COL_FIRSTNAME & currentRow).value))
                    .fullNameCleaned = ubertNH_Utils.GetCleanedFullName(.lastName, .firstName)
                    .subject = Trim(CStr(srcSheet.Range(ubertNH_Config.SRC_COL_SUBJECT & currentRow).value))
                    .value = aqValue
                End With
                
                ' Add to collection
                records.Add record
            End If
        End If
    Next currentRow
    
    Set ReadSourceRecords = records
End Function

' Read date from Kinder sheet
Public Function ReadTransferDate(ByVal srcWorkbook As Workbook, ByRef outMonth As Integer, ByRef outSemester As Integer) As Boolean
    Dim dateSheet As Worksheet
    Dim dateValue As Variant
    Dim dateDate As Date
    
    On Error GoTo ErrorHandler
    
    ' Find Kinder sheet
    Set dateSheet = srcWorkbook.Worksheets(ubertNH_Config.DATE_SHEET_NAME)
    
    ' Read date from T2
    dateValue = dateSheet.Range(ubertNH_Config.DATE_CELL).value
    
    ' Check if empty
    If IsEmpty(dateValue) Or Trim(CStr(dateValue)) = "" Then
        MsgBox ubertNH_Config.MSG_DATE_CELL_EMPTY, vbExclamation
        ReadTransferDate = False
        Exit Function
    End If
    
    ' Try to convert to date
    If IsDate(dateValue) Then
        dateDate = CDate(dateValue)
    Else
        MsgBox ubertNH_Config.MSG_INVALID_DATE, vbExclamation
        ReadTransferDate = False
        Exit Function
    End If
    
    ' Extract month
    outMonth = Month(dateDate)
    
    ' Determine semester
    outSemester = ubertNH_Utils.GetSemester(outMonth)
    
    ReadTransferDate = True
    Exit Function
    
ErrorHandler:
    If Err.Number = 9 Then ' Subscript out of range
        MsgBox ubertNH_Config.MSG_DATE_SHEET_NOT_FOUND, vbExclamation
    Else
        MsgBox "Fehler beim Lesen des Datums: " & Err.Description, vbExclamation
    End If
    ReadTransferDate = False
End Function
