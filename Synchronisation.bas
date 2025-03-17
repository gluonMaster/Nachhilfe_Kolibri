Attribute VB_Name = "Synchronisation"
Option Explicit

Sub SynchronizeFromCopies()
'    On Error GoTo ErrorHandler
    
    ' Declare necessary variables
    Dim originalWb As Workbook
    Dim originalWs As Worksheet
    Dim backupPath As String
    Dim backupFileName As String
    Dim activeSheetName As String
    Dim originalPath As String
    Dim originalFileName As String
    Dim employees As Variant
    Dim employee As Variant
    Dim copyFileName As String
    Dim copyFullPath As String
    Dim copyWb As Workbook
    Dim copyWs As Worksheet
    Dim lastRowCopy As Long
    Dim lastRowOriginal As Long
    Dim dict As Object
    Dim key As String
    Dim i As Long
    Dim j As Long
    Dim hasNonZeroJtoAN As Boolean
    Dim existingRow As Long
    Dim employeeName As String
    Dim fullKey As String
    Dim dictKeys As Object
    Dim dictFullRecords As Object
    Dim fd As FileDialog
    Dim selectedFolder As String
    
    ' Initialize variables
    Set originalWb = ThisWorkbook
    Set originalWs = originalWb.ActiveSheet ' Change if a specific sheet is needed
    activeSheetName = ActiveSheet.name

    ' Step 1: Create a backup copy of the original workbook
    originalPath = originalWb.Path
    originalFileName = originalWb.name
    backupFileName = Left(originalFileName, InStrRev(originalFileName, ".") - 1) & "_old" & mid(originalFileName, InStrRev(originalFileName, "."))
    backupPath = originalPath & "\" & backupFileName

    ' Check if backup already exists to avoid overwriting
    If Dir(backupPath) <> "" Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Backup file '" & backupFileName & "' already exists in the directory. Do you want to create an additional backup?", vbYesNo + vbQuestion, "Backup Exists")
        
        If response = vbYes Then
            ' Find the next available backup filename with a numeric suffix
            Dim counter As Integer
            counter = 1
            Do While Dir(originalPath & "\" & Left(originalFileName, InStrRev(originalFileName, ".") - 1) & "_old" & counter & mid(originalFileName, InStrRev(originalFileName, "."))) <> ""
                counter = counter + 1
            Loop
            
            ' Set the backup path with the new suffix
            backupPath = originalPath & "\" & Left(originalFileName, InStrRev(originalFileName, ".") - 1) & "_old" & counter & mid(originalFileName, InStrRev(originalFileName, "."))
            originalWb.SaveCopyAs fileName:=backupPath
            MsgBox "Additional backup created as '" & backupPath & "'", vbInformation, "Backup Created"
            
        Else
            ' Ask if the user wants to overwrite the existing backup
            response = MsgBox("The existing backup file will be overwritten. Are you sure?", vbYesNo + vbExclamation, "Overwrite Backup?")
            If response = vbNo Then
                MsgBox "Operation cancelled by the user.", vbInformation, "Cancelled"
                Exit Sub
            Else
                ' Overwrite the existing backup
                originalWb.SaveCopyAs fileName:=backupPath
                MsgBox "Backup file '" & backupFileName & "' has been overwritten.", vbInformation, "Backup Overwritten"
            End If
        End If
    Else
        ' No existing backup, so create a new one
        originalWb.SaveCopyAs fileName:=backupPath
        MsgBox "Backup file created as '" & backupFileName & "'", vbInformation, "Backup Created"
    End If
    
    ' Step 2: Define the list of employees
    employees = Array("Nadia", "Valentina", "Anna")
    
    ' Step 3: Load existing records into dictionaries for efficient duplicate checking
    Set dictKeys = CreateObject("Scripting.Dictionary")
    Set dictFullRecords = CreateObject("Scripting.Dictionary")
    
    lastRowOriginal = originalWs.Cells(originalWs.Rows.Count, "A").End(xlUp).row
    
    If lastRowOriginal >= 11 Then
        For i = 11 To lastRowOriginal
            ' Create a key based on columns A to H
            key = ""
            For j = 1 To 8 ' Columns A (1) to H (8)
                key = key & "|" & Trim(CStr(originalWs.Cells(i, j).value))
            Next j
            ' Create a full key based on columns A to H and J to AN
            fullKey = key
            For j = 10 To 40 ' Columns J (10) to AN (40)
                fullKey = fullKey & "|" & Trim(CStr(originalWs.Cells(i, j).value))
            Next j
            ' Add keys to dictionaries
            If Not dictKeys.exists(key) Then dictKeys.Add key, i
            If Not dictFullRecords.exists(fullKey) Then dictFullRecords.Add fullKey, i
        Next i
    End If
    
    ' **Prompt the user to select the destination folder**
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .title = "Select Destination Folder for File Copies"
        .AllowMultiSelect = False
        If .Show = -1 Then ' If the user selects a folder
            selectedFolder = .SelectedItems(1)
        Else
            ' If the user cancels the dialog, exit the macro
            MsgBox "Operation cancelled by the user.", vbInformation, "Cancelled"
            Exit Sub
        End If
    End With
    
    ' Step 4: Improve performance by disabling screen updating and automatic calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Step 5: Process each copy file and add the records
    ' Initialize error list collection
    Dim errorList As Collection
    Set errorList = New Collection

    ' Initialize or clear the ErrorLog worksheet
    Dim errorLogWs As Worksheet
    On Error Resume Next
    Set errorLogWs = originalWb.Worksheets("ErrorLog")
    On Error GoTo 0

    If errorLogWs Is Nothing Then
        Set errorLogWs = originalWb.Worksheets.Add(After:=originalWb.Worksheets(originalWb.Worksheets.Count))
        errorLogWs.name = "ErrorLog"
    Else
        errorLogWs.Cells.Clear
    End If

    ' Initialize ErrorLog headers
    With errorLogWs
        .Range("A1").value = "File Name"
        .Range("B1").value = "Row Number"
        .Range("C1").value = "Reason"
        .Range("A1:C1").Font.Bold = True
    End With

    ' Process each copy file
    For Each employee In employees
        ' Determine the name of the copy file
        copyFileName = Left(originalFileName, InStrRev(originalFileName, ".") - 1) & "_" & employee & ".xlsm"
        copyFullPath = selectedFolder & "\" & copyFileName
        
        ' Check if the copy file exists
        If Dir(copyFullPath) = "" Then
            MsgBox "The copy file '" & copyFileName & "' does not exist in the directory.", vbExclamation, "File Not Found"
            GoTo RestoreSettings
        End If
        
        ' Extract the employee name from the copy file name
        employeeName = employee ' The variable 'employee' directly contains the name from the array
        
        ' Open the copy workbook
        Set copyWb = Workbooks.Open(copyFullPath)
        Set copyWs = copyWb.Sheets(activeSheetName) ' Use the same sheet
        
        ' Find the last used row in the copy workbook
        lastRowCopy = copyWs.Cells(copyWs.Rows.Count, "A").End(xlUp).row
        
        ' Loop through each record in the copy workbook starting from row 11
        For i = 11 To lastRowCopy
            ' Create keys
            key = ""
            For j = 1 To 8 ' Columns A (1) to H (8)
                key = key & "|" & Trim(CStr(copyWs.Cells(i, j).value))
            Next j
            fullKey = key
            For j = 10 To 40 ' Columns J (10) to AN (40)
                fullKey = fullKey & "|" & Trim(CStr(copyWs.Cells(i, j).value))
            Next j
            
            ' Validate required cells
            Dim cellE As Variant, cellAU As Variant, cellAV As Variant
            cellE = copyWs.Cells(i, "E").value
            cellAU = copyWs.Cells(i, "AU").value
            cellAV = copyWs.Cells(i, "AV").value
            
            Dim isValid As Boolean
            isValid = True ' Assume valid unless a check fails
            
            Dim reason As String
            reason = ""
            
            ' Check if E, AU, or AV are empty
            If IsEmpty(cellE) Or IsEmpty(cellAU) Or IsEmpty(cellAV) Then
                isValid = False
                reason = "One or more required cells (E, AU, AV) are empty."
            Else
                ' Check if E contains "I" or "G" (specific criteria)
                If Not (cellE = "I" Or cellE = "G") Then
                    isValid = False
                    reason = "Cell E does not contain 'I' or 'G'."
                End If
            End If
            
            If Not isValid Then
                ' Add invalid record to the error list
                errorList.Add Array(copyFileName, i, reason)
                ' Skip to the next record without processing
                GoTo NextRecord
            End If
            
            ' Check if the full record already exists
            If dictFullRecords.exists(fullKey) Then
                ' Exact record already exists; do nothing
            Else
                If dictKeys.exists(key) Then
                    ' A record with the same key exists
                    existingRow = dictKeys(key)
                    
                    ' Check if columns J to AN are filled in the existing and incoming records
                    Dim existingHasJtoAN As Boolean, copyHasJtoAN As Boolean
                    existingHasJtoAN = False
                    copyHasJtoAN = False
                    
                    For j = 10 To 40
                        If Not IsEmpty(originalWs.Cells(existingRow, j).value) And originalWs.Cells(existingRow, j).value <> 0 Then
                            existingHasJtoAN = True
                            Exit For
                        End If
                    Next j
                    For j = 10 To 40
                        If Not IsEmpty(copyWs.Cells(i, j).value) And copyWs.Cells(i, j).value <> 0 Then
                            copyHasJtoAN = True
                            Exit For
                        End If
                    Next j
                    
                    If Not existingHasJtoAN And copyHasJtoAN Then
                        ' Overwrite the existing record with data from the incoming record
                        For j = 1 To 40
                            originalWs.Cells(existingRow, j).value = copyWs.Cells(i, j).value
                        Next j
                        ' Copy AU and AV
                        originalWs.Cells(existingRow, "AU").value = copyWs.Cells(i, "AU").value
                        originalWs.Cells(existingRow, "AV").value = copyWs.Cells(i, "AV").value
                        ' Set AY to the employee name
                        originalWs.Cells(existingRow, "AY").value = employeeName
                        ' Update fullKey in dictFullRecords
                        ' First delete the old fullKey associated with the existing row
                        Dim dictKeyItem As Variant
                        For Each dictKeyItem In dictFullRecords.keys
                            If dictFullRecords(dictKeyItem) = existingRow Then
                                dictFullRecords.Remove dictKeyItem
                                Exit For
                            End If
                        Next dictKeyItem
                        ' Redefine the fullKey value
                        fullKey = key
                        For j = 10 To 40
                            fullKey = fullKey & "|" & Trim(CStr(originalWs.Cells(existingRow, j).value))
                        Next j
                        ' Add the updated fullKey to dictFullRecords
                        dictFullRecords.Add fullKey, existingRow
                    ElseIf existingHasJtoAN And copyHasJtoAN Then
                        ' Both have filled columns J to AN
                        ' Check if data differs
                        Dim isDifferent As Boolean
                        isDifferent = False
                        For j = 10 To 40
                            If Trim(CStr(originalWs.Cells(existingRow, j).value)) <> Trim(CStr(copyWs.Cells(i, j).value)) Then
                                isDifferent = True
                                Exit For
                            End If
                        Next j
                        If isDifferent Then
                            ' Add the incoming record as new
                            lastRowOriginal = lastRowOriginal + 1
                            For j = 1 To 40
                                originalWs.Cells(lastRowOriginal, j).value = copyWs.Cells(i, j).value
                            Next j
                            ' Copy AU and AV
                            originalWs.Cells(lastRowOriginal, "AU").value = copyWs.Cells(i, "AU").value
                            originalWs.Cells(lastRowOriginal, "AV").value = copyWs.Cells(i, "AV").value
                            ' Set AY to the employee name
                            originalWs.Cells(lastRowOriginal, "AY").value = employeeName
                            ' Add new fullKey to dictFullRecords
                            dictFullRecords.Add fullKey, lastRowOriginal
                        Else
                            ' Data matches; do nothing
                        End If
                    Else
                        ' Do nothing in other cases
                        ' If the existing record has filled J-AN and the incoming does not, or both are empty
                    End If
                Else
                    ' A record with the same key does not exist; add as new
                    lastRowOriginal = lastRowOriginal + 1
                    For j = 1 To 40
                        originalWs.Cells(lastRowOriginal, j).value = copyWs.Cells(i, j).value
                    Next j
                    ' Copy AU and AV
                    originalWs.Cells(lastRowOriginal, "AU").value = copyWs.Cells(i, "AU").value
                    originalWs.Cells(lastRowOriginal, "AV").value = copyWs.Cells(i, "AV").value
                    ' Set AY to the employee name
                    originalWs.Cells(lastRowOriginal, "AY").value = employeeName
                    ' Add keys to dictionaries
                    dictKeys.Add key, lastRowOriginal
                    dictFullRecords.Add fullKey, lastRowOriginal
                End If
            End If
            
NextRecord:
        Next i
        
        ' Close the copy workbook without saving changes
        copyWb.Close SaveChanges:=False
    Next employee

    ' After processing all files, write any errors to the ErrorLog
    If errorList.Count > 0 Then
        Dim errorRow As Long
        errorRow = 2 ' Start from the second row, as first row has headers
        
        Dim errorItem As Variant
        For Each errorItem In errorList
            errorLogWs.Cells(errorRow, "A").value = errorItem(0) ' File Name
            errorLogWs.Cells(errorRow, "B").value = errorItem(1) ' Row Number
            errorLogWs.Cells(errorRow, "C").value = errorItem(2) ' Reason
            errorRow = errorRow + 1
        Next errorItem
        
        ' Auto-fit the columns in ErrorLog
        errorLogWs.Columns("A:C").AutoFit
    End If

    ' Step 6: Sort the original workbook's sheet based on column B in ascending order
    lastRowOriginal = originalWs.Cells(originalWs.Rows.Count, "A").End(xlUp).row
    If lastRowOriginal >= 11 Then
        originalWs.Range("A11:AY" & lastRowOriginal).Sort _
            Key1:=originalWs.Range("B11"), Order1:=xlAscending, _
            Key2:=originalWs.Range("C11"), Order1:=xlAscending, _
            Header:=xlNo
    End If
    
    ' Step 6.1: Highlight duplicate records based on columns A through H
    Dim duplicateDict As Object
    Dim duplicateKey As String
    Dim dupRow As Long
    
    Set duplicateDict = CreateObject("Scripting.Dictionary")
    
    For i = 11 To lastRowOriginal
        ' Create a unique key based on the values of columns A through H
        duplicateKey = ""
        For j = 1 To 8
            duplicateKey = duplicateKey & "|" & Trim(CStr(originalWs.Cells(i, j).value))
        Next j
        
        If duplicateDict.exists(duplicateKey) Then
            ' Get the row number of an existing record
            dupRow = duplicateDict(duplicateKey)
            ' Highlight cells A through H in the existing row
            originalWs.Range(originalWs.Cells(dupRow, 1), originalWs.Cells(dupRow, 8)).Interior.Color = RGB(173, 216, 230)
            ' Highlight cells A through H in the current duplicate row
            originalWs.Range(originalWs.Cells(i, 1), originalWs.Cells(i, 8)).Interior.Color = RGB(173, 216, 230)
        Else
            ' Add the key and row number to the dictionary
            duplicateDict.Add duplicateKey, i
        End If
    Next i

    ' Step 7: Restore settings and inform the user
RestoreSettings:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "The first stage of synchronization is completed. Academic workload data has been synchronized successfully. Wait until the processing of the list of children is complete.", vbInformation, "Success"

    ' Call the Kinder synchronization
    Call SynchronizeKinder(employees, selectedFolder)
    Exit Sub

'ErrorHandler:
'    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
    
End Sub


Sub SynchronizeKinder(employees As Variant, selectedFolder As String)
    On Error GoTo ErrorHandler
    
    ThisWorkbook.IsMacroRunning = True

    ' Improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Declare necessary variables
    Dim originalWb As Workbook
    Dim originalWs As Worksheet
    Dim backupFileName As String
    Dim copyFileName As String
    Dim copyFullPath As String
    Dim copyWb As Workbook
    Dim copyWs As Worksheet
    Dim lastRowCopy As Long
    Dim lastRowOriginal As Long
    Dim dictOriginal As Object
    Dim dictAdded As Object
    Dim key As String
    Dim i As Long
    Dim j As Long
    Dim hasNonZeroA As Boolean
    Dim hasDuplicates As Boolean
    Dim recordNumber As Variant
    Dim employee As Variant
    Dim operatorName As String
    Dim maxRecordOriginal As Long
    Dim uniqueKey As String
    Dim rowsToDelete As Collection
    Dim rowIndex As Variant
    Dim keyNum As Variant

    ' Initialize variables
    Set originalWb = ThisWorkbook
    Set originalWs = originalWb.Sheets("Kinder") ' Ensure the sheet name is exactly "Kinder"

    ' Initialize dictionaries
    Set dictOriginal = CreateObject("Scripting.Dictionary")
    Set dictAdded = CreateObject("Scripting.Dictionary")
    Set rowsToDelete = New Collection

    ' Load existing records from original "Kinder" sheet into dictOriginal
    lastRowOriginal = originalWs.Cells(originalWs.Rows.Count, "A").End(xlUp).row

    If lastRowOriginal >= 5 Then
        For i = 5 To lastRowOriginal
            recordNumber = originalWs.Cells(i, "A").value
            If Not IsEmpty(recordNumber) Then
                If Not dictOriginal.exists(recordNumber) Then
                    Set dictOriginal(recordNumber) = New Collection
                End If
                dictOriginal(recordNumber).Add i
            End If
        Next i
    End If

    ' Process each employee copy
    For Each employee In employees
        ' Define the copy file name
        copyFileName = Left(originalWb.name, InStrRev(originalWb.name, ".") - 1) & "_" & employee & ".xlsm"
        copyFullPath = selectedFolder & "\" & copyFileName

        ' Check if the copy file exists
        If Dir(copyFullPath) = "" Then
            MsgBox "Copy file '" & copyFileName & "' does not exist in the directory.", vbExclamation, "File Not Found"
            GoTo NextEmployee
        End If

        ' Extract operator name from file name
        operatorName = employee ' Assuming the employee name is the part after the underscore

        ' Open the copy workbook
        Set copyWb = Workbooks.Open(copyFullPath)
        Set copyWs = copyWb.Sheets("Kinder") ' Ensure the sheet name is exactly "Kinder"

        ' Validation - Check for non-zero and unique values in Column A
        hasNonZeroA = True
        hasDuplicates = False
        Dim seenNumbers As Object
        Set seenNumbers = CreateObject("Scripting.Dictionary")

        lastRowCopy = copyWs.Cells(copyWs.Rows.Count, "A").End(xlUp).row

        For i = 5 To lastRowCopy
            recordNumber = copyWs.Cells(i, "A").value
            If IsEmpty(recordNumber) Or recordNumber = 0 Then
                hasNonZeroA = False
                Exit For
            End If
            If seenNumbers.exists(recordNumber) Then
                hasDuplicates = True
                Exit For
            Else
                seenNumbers.Add recordNumber, True
            End If
        Next i

        If Not hasNonZeroA Then
            MsgBox "Error: Found empty or zero record numbers in 'Kinder' sheet of '" & copyFileName & "'. Synchronization aborted for this file.", vbCritical, "Validation Error"
            copyWb.Close SaveChanges:=False
            GoTo NextEmployee
        End If

        If hasDuplicates Then
            MsgBox "Error: Found duplicate record numbers in 'Kinder' sheet of '" & copyFileName & "'. Synchronization aborted for this file.", vbCritical, "Validation Error"
            copyWb.Close SaveChanges:=False
            GoTo NextEmployee
        End If

        ' Sort the copy's "Kinder" sheet by Column A ascending
        copyWs.Range("A5:T" & lastRowCopy).Sort _
            Key1:=copyWs.Range("A5"), Order1:=xlAscending, _
            Header:=xlNo

        ' Identify missing record numbers
        Dim expectedNumber As Long
        Dim currentNumber As Long
        Dim missingList As Collection
        Set missingList = New Collection
        expectedNumber = copyWs.Cells(5, "A").value

        For i = 5 To lastRowCopy
            currentNumber = copyWs.Cells(i, "A").value
            While expectedNumber < currentNumber
                missingList.Add expectedNumber
                expectedNumber = expectedNumber + 1
            Wend
            expectedNumber = currentNumber + 1
        Next i

        ' Collect rows to delete from original workbook
        If missingList.Count > 0 Then
            For Each recordNumber In missingList
                If dictOriginal.exists(recordNumber) Then
                    For Each rowIndex In dictOriginal(recordNumber)
                        rowsToDelete.Add rowIndex
                    Next rowIndex
                End If
            Next recordNumber
        End If

        ' Delete rows in descending order to prevent shifting issues
        If rowsToDelete.Count > 0 Then
            ' Sort rowsToDelete in descending order
            Dim sortedRows() As Long
            ReDim sortedRows(1 To rowsToDelete.Count)
            For i = 1 To rowsToDelete.Count
                sortedRows(i) = rowsToDelete(i)
            Next i
            ' Sort the array in descending order
            Call QuickSortDescending(sortedRows, LBound(sortedRows), UBound(sortedRows))
            ' Delete the rows
            For i = 1 To UBound(sortedRows)
                originalWs.Rows(sortedRows(i)).Delete
            Next i
            ' Clear rowsToDelete for next use
            Set rowsToDelete = New Collection
            ' Rebuild dictOriginal after deletions
            Set dictOriginal = CreateObject("Scripting.Dictionary")
            lastRowOriginal = originalWs.Cells(originalWs.Rows.Count, "A").End(xlUp).row
            If lastRowOriginal >= 5 Then
                For i = 5 To lastRowOriginal
                    recordNumber = originalWs.Cells(i, "A").value
                    If Not IsEmpty(recordNumber) Then
                        If Not dictOriginal.exists(recordNumber) Then
                            Set dictOriginal(recordNumber) = New Collection
                        End If
                        dictOriginal(recordNumber).Add i
                    End If
                Next i
            End If
        End If

        ' Synchronize records
        For i = 5 To lastRowCopy
            recordNumber = copyWs.Cells(i, "A").value
            If dictOriginal.exists(recordNumber) Then
                ' Iterate through all original rows with this recordNumber
                Dim originalRows As Collection
                Set originalRows = dictOriginal(recordNumber)
                Dim matched As Boolean
                matched = False
                For Each rowIndex In originalRows
                    ' Compare records A-L, N, O
                    Dim isDifferent As Boolean
                    isDifferent = False
                    For j = 1 To 12 ' Columns A (1) to L (12)
                        If originalWs.Cells(rowIndex, j).value <> copyWs.Cells(i, j).value Then
                            isDifferent = True
                            Exit For
                        End If
                    Next j
                    If Not isDifferent Then
                        For j = 14 To 15 ' Columns N (14) to O (15)
                            If originalWs.Cells(rowIndex, j).value <> copyWs.Cells(i, j).value Then
                                isDifferent = True
                                Exit For
                            End If
                        Next j
                    End If

                    If isDifferent Then
                        ' Check Column U in original
                        If IsEmpty(originalWs.Cells(rowIndex, "U").value) Then
                            ' Overwrite the original record with the copy
                            For j = 1 To 12 ' Columns A-L
                                originalWs.Cells(rowIndex, j).value = copyWs.Cells(i, j).value
                            Next j
                            For j = 14 To 15 ' Columns N,O
                                originalWs.Cells(rowIndex, j).value = copyWs.Cells(i, j).value
                            Next j
                            ' Insert operator name in Column U
                            originalWs.Cells(rowIndex, "U").value = operatorName
                            ' Mark as matched to avoid adding as new
                            matched = True
                            Exit For
                        End If
                    Else
                        ' Records are identical; no action needed
                        matched = True
                        Exit For
                    End If
                Next rowIndex

                If Not matched Then
                    ' Either all existing records are marked, or no differences found
                    ' Check if this exact record already exists with operatorName
                    Dim exists As Boolean
                    exists = False
                    For Each rowIndex In originalRows
                        If originalWs.Cells(rowIndex, "U").value = operatorName Then
                            ' Assuming that operatorName in column U indicates this record has been processed by this operator
                            exists = True
                            Exit For
                        End If
                    Next rowIndex

                    If Not exists Then
                        ' Add the copy as a new record
                        ' Ensure that this exact record does not already exist
                        Dim duplicateFound As Boolean
                        duplicateFound = False
                        For Each rowIndex In originalRows
                            Dim isExactMatch As Boolean
                            isExactMatch = True
                            For j = 1 To 12 ' Columns A-L
                                If originalWs.Cells(rowIndex, j).value <> copyWs.Cells(i, j).value Then
                                    isExactMatch = False
                                    Exit For
                                End If
                            Next j
                            If isExactMatch Then
                                For j = 14 To 15 ' Columns N-O
                                    If originalWs.Cells(rowIndex, j).value <> copyWs.Cells(i, j).value Then
                                        isExactMatch = False
                                        Exit For
                                    End If
                                Next j
                            End If
                            If isExactMatch Then
                                duplicateFound = True
                                Exit For
                            End If
                        Next rowIndex

                        If Not duplicateFound Then
                            lastRowOriginal = lastRowOriginal + 1
                            For j = 1 To 12 ' Columns A-L
                                originalWs.Cells(lastRowOriginal, j).value = copyWs.Cells(i, j).value
                            Next j
                            For j = 14 To 15 ' Columns N-O
                                originalWs.Cells(lastRowOriginal, j).value = copyWs.Cells(i, j).value
                            Next j
                            ' Insert operator name in Column U
                            originalWs.Cells(lastRowOriginal, "U").value = operatorName
                            ' Add to dictionary
                            If Not dictOriginal.exists(recordNumber) Then
                                Set dictOriginal(recordNumber) = New Collection
                            End If
                            dictOriginal(recordNumber).Add lastRowOriginal
                        End If
                    End If
                End If
            Else
                ' Record does not exist in original, check if it should be appended
                ' Find the maximum record number in original
                If lastRowOriginal < 5 Then
                    maxRecordOriginal = 0
                Else
                    maxRecordOriginal = originalWs.Cells(originalWs.Rows.Count, "A").End(xlUp).value
                End If

                If recordNumber > maxRecordOriginal Then
                    ' Append the record to original
                    lastRowOriginal = lastRowOriginal + 1
                    For j = 1 To 12 ' Columns A-L
                        originalWs.Cells(lastRowOriginal, j).value = copyWs.Cells(i, j).value
                    Next j
                    For j = 14 To 15 ' Columns N-O
                        originalWs.Cells(lastRowOriginal, j).value = copyWs.Cells(i, j).value
                    Next j
                    ' Insert operator name in Column U
                    originalWs.Cells(lastRowOriginal, "U").value = operatorName
                    ' Add to dictionary
                    If Not dictOriginal.exists(recordNumber) Then
                        Set dictOriginal(recordNumber) = New Collection
                    End If
                    dictOriginal(recordNumber).Add lastRowOriginal
                End If
            End If
        Next i

        ' Close the copy workbook without saving changes
        copyWb.Close SaveChanges:=False

NextEmployee:
    Next employee

    ' Sort the original workbook's "Kinder" sheet by Column C and D ascending (A-U)
    lastRowOriginal = originalWs.Cells(originalWs.Rows.Count, "A").End(xlUp).row
    If lastRowOriginal >= 5 Then
        originalWs.Range("A5:U" & lastRowOriginal).Sort _
            Key1:=originalWs.Range("C5"), Order1:=xlAscending, _
            Key2:=originalWs.Range("D5"), Order2:=xlAscending, _
            Header:=xlNo
    End If

    ' Inform the user of successful synchronization
    MsgBox "Synchronization is completely finished! Kinder sheets synchronization completed successfully.", vbInformation, "Success"

    ' Restore Excel settings
RestoreSettings:
    ThisWorkbook.IsMacroRunning = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    ThisWorkbook.IsMacroRunning = False
    MsgBox "An error occurred during Kinder synchronization: " & Err.Description, vbCritical, "Error"
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' Helper Subroutine: QuickSort in Descending Order
Sub QuickSortDescending(arr() As Long, first As Long, last As Long)
    Dim low As Long, high As Long
    Dim mid As Long
    Dim temp As Long

    low = first
    high = last
    mid = arr((first + last) \ 2)

    Do While low <= high
        Do While arr(low) > mid
            low = low + 1
        Loop
        Do While arr(high) < mid
            high = high - 1
        Loop
        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then QuickSortDescending arr, first, high
    If low < last Then QuickSortDescending arr, low, last
End Sub


