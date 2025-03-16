Attribute VB_Name = "Archivierung"
Option Explicit

' Function at the module level
Function IsDateRangeActive(startDate As Variant, endDate As Variant, referenceMonth As Integer, referenceYear As Integer) As Boolean
    IsDateRangeActive = False
    
    If IsDate(startDate) And IsDate(endDate) Then
        ' 1. If the date range spans the month of the reference date
        If (Year(startDate) = referenceYear And Month(startDate) = referenceMonth) Or _
           (Year(endDate) = referenceYear And Month(endDate) = referenceMonth) Or _
           (startDate <= DateSerial(referenceYear, referenceMonth, 1) And _
            endDate >= DateSerial(referenceYear, referenceMonth, 1)) Then
            IsDateRangeActive = True
        ' 2. If the date range is in a future month relative to the reference date
        ElseIf (Year(startDate) > referenceYear) Or _
               (Year(startDate) = referenceYear And Month(startDate) > referenceMonth) Then
            IsDateRangeActive = True
        End If
    End If
End Function

Sub ArchiveAndCleanKinder_Optimized_Fixed()
    Dim wsKinder As Worksheet
    Dim wsArchiv As Worksheet
    Dim lastRowKinder As Long
    Dim lastRowArchiv As Long
    Dim i As Long, j As Long
    Dim todayDate As Date
    Dim dictArchiv As Object
    Dim arrArchiv As Variant
    Dim arrKinder As Variant
    Dim collToArchive As Collection
    Dim archiveCount As Long
    Dim keepCount As Long
    Dim arrKeep As Variant
    Dim key As String
    Dim startDate As Variant
    Dim endDate As Variant
    Dim archiveStartRow As Long
    Dim arrToArchive As Variant
    Dim serials As Variant
    Dim collToKeep As Collection ' Collection for records to keep
    Dim referenceDate As Date
    Dim referenceMonth As Integer
    Dim referenceYear As Integer
    Dim collActiveRecords As Collection
    Dim collInactiveRecords As Collection
    Dim activeCount As Long
    Dim inactiveCount As Long
    Dim isActive As Boolean
    Dim totalRecords As Long
    Dim wasInactive As Boolean
    Dim reconsideredCount As Long
    Dim combinedRecords As Variant
    Dim response As VbMsgBoxResult
    
    ThisWorkbook.IsMacroRunning = True
    
    ' Disable screen updating, automatic calculations, and events to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo Cleanup ' Ensure that settings are restored even if an error occurs
    
    ' Set references to the worksheets
    Set wsKinder = ThisWorkbook.Sheets("Kinder")
    Set wsArchiv = ThisWorkbook.Sheets("Archiv")
    
    todayDate = Date
    
    ' Determine the last filled row on the Kinder sheet
    lastRowKinder = wsKinder.Cells(wsKinder.Rows.Count, "A").End(xlUp).row
    
    ' Define the range of data
    Dim dataRange As Range
    Set dataRange = wsKinder.Range("A5:V" & lastRowKinder)
    
    ' Clear previous formatting and comments
    dataRange.Interior.ColorIndex = xlNone
    dataRange.Columns("V").ClearContents
    dataRange.Columns("U").ClearContents
    Dim cell As Range
    For Each cell In dataRange
        cell.ClearComments
    Next cell
    
    ' Determine the last filled row on the Archiv sheet
    lastRowArchiv = wsArchiv.Cells(wsArchiv.Rows.Count, "A").End(xlUp).row
    
    ' Initialize Dictionary to store keys from Archiv
    Set dictArchiv = CreateObject("Scripting.Dictionary")
    dictArchiv.CompareMode = vbTextCompare ' Case-insensitive comparison
    
    ' Load data from Archiv into an array and populate the Dictionary
    If lastRowArchiv >= 5 Then
        arrArchiv = wsArchiv.Range("B5:V" & lastRowArchiv).value  ' Expand range to include columns S and T
        For i = 1 To UBound(arrArchiv, 1)
            key = ""
            ' Add B-D (indices 1-3 in the array)
            For j = 1 To 3
                key = key & Trim(LCase(arrArchiv(i, j)))
            Next j
            
            ' Add K-O (indices 10-14 in the array)
            For j = 10 To 14
                key = key & Trim(LCase(arrArchiv(i, j)))
            Next j
            
            ' Add S and T (indices 18-19 in the array)
            key = key & Trim(LCase(arrArchiv(i, 18))) & Trim(LCase(arrArchiv(i, 19)))
            
            If Not dictArchiv.exists(key) Then
                dictArchiv.Add key, True
            End If
        Next i
    End If
    
    ' Load data from Kinder into an array
    If lastRowKinder >= 5 Then
        arrKinder = wsKinder.Range("A5:V" & lastRowKinder).value
    Else
        ReDim arrKinder(1 To 1, 1 To 22) ' Initialize as 1x22 array to prevent errors
    End If
    
    ' Get the reference date from cell T2 of the Kinder sheet
    On Error Resume Next
    referenceDate = wsKinder.Range("T2").value
    On Error GoTo 0
    
    ' If reference date is not valid, use today's date
    If Not IsDate(referenceDate) Then
        referenceDate = Date
    End If
    
    ' Get the month and year of the reference date
    referenceMonth = Month(referenceDate)
    referenceYear = Year(referenceDate)
    
    ' Initialize Collections for active and inactive records
    Set collActiveRecords = New Collection
    Set collInactiveRecords = New Collection
    Set collToArchive = New Collection
    archiveCount = 0
    activeCount = 0
    inactiveCount = 0
    reconsideredCount = 0
    
    ' First pass - check existing data and determine active status
    wasInactive = False
    
    ' Check existing data for inactive records that might need reactivation
    If lastRowKinder >= 5 Then
        ' Scan existing rows for gray font color (inactive records)
        For i = 5 To lastRowKinder
            Dim startDateCell As Range
            Dim endDateCell As Range
            
            Set startDateCell = wsKinder.Cells(i, "G")
            Set endDateCell = wsKinder.Cells(i, "H")
            
            If wsKinder.Cells(i, "A").Font.ColorIndex = 15 Then ' This was an inactive record
                wasInactive = True
                
                ' Check if the date range is now active
                If IsDateRangeActive(startDateCell.value, endDateCell.value, referenceMonth, referenceYear) Then
                    ' Reactivate this record
                    wsKinder.Range("A" & i & ":V" & i).Font.ColorIndex = xlAutomatic ' Black font
                    reconsideredCount = reconsideredCount + 1
                End If
            End If
        Next i
        
        ' Inform user if records were reactivated
        If reconsideredCount > 0 Then
            MsgBox reconsideredCount & " previously inactive records have been reactivated based on new date criteria.", vbInformation
        End If
    End If
    
    ' Iterate through each record in Kinder
    For i = 1 To UBound(arrKinder, 1)
        ' Process key for archive comparison
        key = ""
        ' Add B-D (indices 2-4 in the arrKinder array)
        For j = 2 To 4
            key = key & Trim(LCase(arrKinder(i, j)))
        Next j
        
        ' Add K-O (indices 11-15 in the arrKinder array)
        For j = 11 To 15
            key = key & Trim(LCase(arrKinder(i, j)))
        Next j
        
        ' Add S and T (indices 19-20 in the arrKinder array)
        key = key & Trim(LCase(arrKinder(i, 19))) & Trim(LCase(arrKinder(i, 20)))
        
        ' Check if the record exists in Archiv
        If Not dictArchiv.exists(key) Then
            ' Record does not exist in Archiv; add to archive collection
            archiveCount = archiveCount + 1
            ' Create a temporary array for the row
            Dim tempRowArchive(1 To 22) As Variant
            For j = 1 To 22
                tempRowArchive(j) = arrKinder(i, j)
            Next j
            ' Add a timestamp in column V (22)
            tempRowArchive(22) = Now
            ' Add the row to the archive collection
            collToArchive.Add tempRowArchive
        End If
        
        ' Check if the record should be active or inactive based on date criteria
        On Error Resume Next
        startDate = arrKinder(i, 7) ' Column G
        endDate = arrKinder(i, 8)   ' Column H
        On Error GoTo 0
        
        ' Use the function to determine if record is active
        isActive = IsDateRangeActive(startDate, endDate, referenceMonth, referenceYear)
        
        ' Create a temporary array for the row to keep
        Dim tempRowKeep(1 To 22) As Variant
        For j = 1 To 22
            tempRowKeep(j) = arrKinder(i, j)
        Next j
        
        ' Add row to the appropriate collection based on active status
        If isActive Then
            activeCount = activeCount + 1
            collActiveRecords.Add tempRowKeep
        Else
            inactiveCount = inactiveCount + 1
            collInactiveRecords.Add tempRowKeep
        End If
    Next i
    
    ' Process archive records
    If archiveCount > 0 Then
        ' Initialize arrToArchive as a 2D array
        ReDim arrToArchive(1 To archiveCount, 1 To 22)
        ' Transfer data from the collection to the array
        For i = 1 To archiveCount
            For j = 1 To 22
                arrToArchive(i, j) = collToArchive(i)(j)
            Next j
        Next i
        ' Find the next available row in Archiv
        archiveStartRow = lastRowArchiv + 1
        ' Write the archive data to Archiv in one operation
        wsArchiv.Range("A" & archiveStartRow).Resize(archiveCount, 22).value = arrToArchive
        
        ' Apply proper formatting to newly added records
        With wsArchiv.Range("A" & archiveStartRow & ":V" & (archiveStartRow + archiveCount - 1))
            ' Reset any existing formatting first
            .Font.Bold = False
            '.HorizontalAlignment = xlGeneral
            
            ' Apply specific formatting
            .Columns("C").Font.Bold = True
            .Columns("K").Font.Bold = True
            .Columns("M").Font.Bold = True
            
            ' Center align specific columns
            .Columns("K").HorizontalAlignment = xlCenter
            .Columns("L").HorizontalAlignment = xlCenter
            .Columns("S").HorizontalAlignment = xlCenter
            .Columns("T").HorizontalAlignment = xlCenter
        End With
    End If
    
    ' Clear existing data and formatting on Kinder sheet
    If lastRowKinder >= 5 Then
        wsKinder.Rows("5:" & lastRowKinder).ClearContents
        wsKinder.Rows("5:" & lastRowKinder).Font.ColorIndex = xlAutomatic ' Reset font color
    End If
    
    ' Combine active and inactive records for writing back to the Kinder sheet
    totalRecords = activeCount + inactiveCount
    
    If totalRecords > 0 Then
        ' Initialize combined array
        ReDim combinedRecords(1 To totalRecords, 1 To 22)
        
        ' First add all active records
        For i = 1 To activeCount
            For j = 1 To 22
                combinedRecords(i, j) = collActiveRecords(i)(j)
            Next j
        Next i
        
        ' Then add all inactive records
        For i = 1 To inactiveCount
            For j = 1 To 22
                combinedRecords(activeCount + i, j) = collInactiveRecords(i)(j)
            Next j
        Next i
        
        ' Write all records back to the Kinder sheet
        wsKinder.Range("A5").Resize(totalRecords, 22).value = combinedRecords
        
        ' Add the formula back to column M for all records
        wsKinder.Range("M5:M" & 4 + totalRecords).FormulaR1C1 = "=DATEDIF(RC[-1], TODAY(), ""Y"")"
        
        ' Set gray font color for inactive records
        If inactiveCount > 0 Then
            wsKinder.Range("A" & 5 + activeCount & ":V" & 4 + totalRecords).Font.ColorIndex = 15 ' Gray color
            
            ' Ask user if they want to delete inactive records
            response = MsgBox("There are " & inactiveCount & " inactive records. Would you like to delete them?", _
                              vbQuestion + vbYesNo, "Delete Inactive Records")
            
            If response = vbYes Then
                ' Delete inactive records
                wsKinder.Rows((5 + activeCount) & ":" & (4 + totalRecords)).Delete
                
                ' Update serial numbers after deletion
                If activeCount > 0 Then
                    ReDim serials(1 To activeCount, 1 To 1)
                    For i = 1 To activeCount
                        serials(i, 1) = i
                    Next i
                    wsKinder.Range("A5").Resize(activeCount, 1).value = serials
                End If
                
                MsgBox inactiveCount & " inactive records have been deleted.", vbInformation
            End If
        End If
        
        ' Sort the Kinder sheet by columns C and D in ascending order
        If activeCount > 0 Then
            With wsKinder.Sort
                .SortFields.Clear
                .SortFields.Add key:=wsKinder.Range("C5:C" & 4 + totalRecords), Order:=xlAscending
                .SortFields.Add key:=wsKinder.Range("D5:D" & 4 + totalRecords), Order:=xlAscending
                .SetRange wsKinder.Range("A5:V" & 4 + totalRecords)
                .Header = xlNo
                .Apply
            End With
        End If
        
        ' Update serial numbers in column A
        If totalRecords > 0 Then
            ReDim serials(1 To totalRecords, 1 To 1)
            For i = 1 To totalRecords
                serials(i, 1) = i
            Next i
            wsKinder.Range("A5").Resize(totalRecords, 1).value = serials
        End If
    End If
    
    ' Sort the Archiv sheet by columns C and D in ascending order
    If lastRowArchiv >= 5 And archiveCount > 0 Then
        lastRowArchiv = wsArchiv.Cells(wsArchiv.Rows.Count, "A").End(xlUp).row
        With wsArchiv.Sort
            .SortFields.Clear
            .SortFields.Add key:=wsArchiv.Range("C5:C" & lastRowArchiv), Order:=xlAscending
            .SortFields.Add key:=wsArchiv.Range("D5:D" & lastRowArchiv), Order:=xlAscending
            .SetRange wsArchiv.Range("A5:V" & lastRowArchiv)
            .Header = xlNo
            .Apply
        End With
    End If
    
    ' Update serial numbers in column A in list Archiv
    If archiveCount > 0 Then
        Dim lastRowArchivUpdated As Long
        Dim totalArchivRecords As Long
        Dim archivSerials As Variant
        
        lastRowArchivUpdated = wsArchiv.Cells(wsArchiv.Rows.Count, "A").End(xlUp).row
        totalArchivRecords = lastRowArchivUpdated - 4 ' Because the data begins from the row 5

        ReDim archivSerials(1 To totalArchivRecords, 1 To 1)
        
        For i = 1 To totalArchivRecords
            archivSerials(i, 1) = i
        Next i

        wsArchiv.Range("A5").Resize(totalArchivRecords, 1).value = archivSerials
    End If
    
    Call RemoveDuplicatesFromArchiv

    MsgBox "Archiving and cleaning completed successfully!", vbInformation

Cleanup:
    ThisWorkbook.IsMacroRunning = False
    ' Re-enable screen updating, automatic calculations, and events
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' Handle any unexpected errors
    If Err.Number <> 0 Then
        MsgBox "An error occurred: " & Err.Description, vbExclamation
    End If
End Sub

Sub RemoveDuplicatesFromArchiv()
    ' This procedure removes duplicate records from the "Archiv" sheet
    ' Duplicates are defined as records having the same content in cells B, C, D
    ' Only the most recent record (based on timestamp in column V) is kept
    
    Dim wsArchiv As Worksheet
    Dim lastRowArchiv As Long
    Dim i As Long, j As Long
    Dim dictUnique As Object
    Dim keyBCD As String
    Dim mostRecentDate As Date
    Dim archivData As Variant
    Dim rowsToDelete As Collection
    Dim rowIndex As Variant
    Dim rowsDeleted As Long
    
    ' Set reference to the Archiv worksheet
    Set wsArchiv = ThisWorkbook.Sheets("Archiv")
    
    ' Determine the last filled row on the Archiv sheet
    lastRowArchiv = wsArchiv.Cells(wsArchiv.Rows.Count, "A").End(xlUp).row
    
    ' If there are no records or just a header, exit
    If lastRowArchiv < 5 Then Exit Sub
    
    ' Load data from Archiv into an array for better performance
    archivData = wsArchiv.Range("A5:V" & lastRowArchiv).value
    
    ' Create a dictionary to keep track of unique records
    Set dictUnique = CreateObject("Scripting.Dictionary")
    dictUnique.CompareMode = vbTextCompare ' Case-insensitive comparison
    
    ' Create a collection to store rows that should be deleted
    Set rowsToDelete = New Collection
    
    ' First pass: Identify duplicates and determine which to keep
    For i = 1 To UBound(archivData, 1)
        ' Create a key from columns B, C, D (indices 2, 3, 4 in the array)
        keyBCD = Trim(LCase(archivData(i, 2))) & "|" & _
                 Trim(LCase(archivData(i, 3))) & "|" & _
                 Trim(LCase(archivData(i, 4)))
        
        ' Get the timestamp from column V (index 22 in the array)
        Dim currentTimestamp As Variant
        currentTimestamp = archivData(i, 22)
        
        ' Check if this key already exists in our dictionary
        If dictUnique.exists(keyBCD) Then
            ' Compare timestamps to determine which record to keep
            Dim existingRow As Long
            Dim existingTimestamp As Variant
            
            existingRow = dictUnique(keyBCD)
            existingTimestamp = archivData(existingRow, 22)
            
            ' If current record is newer (or existing timestamp is invalid)
            If (Not IsDate(existingTimestamp)) Or _
               (IsDate(currentTimestamp) And currentTimestamp > existingTimestamp) Then
                
                ' Mark the older record for deletion
                On Error Resume Next
                rowsToDelete.Add existingRow
                On Error GoTo 0
                
                ' Update dictionary with the row of the newer record
                dictUnique(keyBCD) = i
            Else
                ' Current record is older or has invalid timestamp, mark it for deletion
                On Error Resume Next
                rowsToDelete.Add i
                On Error GoTo 0
            End If
        Else
            ' This is a unique record, add it to the dictionary
            dictUnique.Add keyBCD, i
        End If
    Next i
    
    ' Sort the row indices in descending order to avoid shifting issues when deleting
    Dim sortedRowsToDelete() As Long
    Dim sortedIndex As Long
    
    If rowsToDelete.Count > 0 Then
        ReDim sortedRowsToDelete(1 To rowsToDelete.Count)
        
        sortedIndex = 1
        For Each rowIndex In rowsToDelete
            sortedRowsToDelete(sortedIndex) = rowIndex
            sortedIndex = sortedIndex + 1
        Next rowIndex
        
        ' Simple bubble sort in descending order
        Dim temp As Long
        Dim swapped As Boolean
        
        Do
            swapped = False
            For i = 1 To UBound(sortedRowsToDelete) - 1
                If sortedRowsToDelete(i) < sortedRowsToDelete(i + 1) Then
                    temp = sortedRowsToDelete(i)
                    sortedRowsToDelete(i) = sortedRowsToDelete(i + 1)
                    sortedRowsToDelete(i + 1) = temp
                    swapped = True
                End If
            Next i
        Loop Until Not swapped
        
        ' Delete the duplicate rows (from bottom to top to avoid shifting issues)
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        
        For i = 1 To UBound(sortedRowsToDelete)
            wsArchiv.Rows(sortedRowsToDelete(i) + 4).Delete ' +4 to account for the offset (data starts at row 5)
        Next i
        
        rowsDeleted = UBound(sortedRowsToDelete)
        
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        ' Update serial numbers in column A
        lastRowArchiv = wsArchiv.Cells(wsArchiv.Rows.Count, "A").End(xlUp).row
        
        If lastRowArchiv >= 5 Then
            Dim totalRows As Long
            Dim serials As Variant
            
            totalRows = lastRowArchiv - 4 ' Data starts at row 5
            
            ReDim serials(1 To totalRows, 1 To 1)
            For i = 1 To totalRows
                serials(i, 1) = i
            Next i
            
            wsArchiv.Range("A5").Resize(totalRows, 1).value = serials
        End If
        
        ' Inform the user
        MsgBox "Removed " & rowsDeleted & " duplicate records from the Archiv sheet.", vbInformation
    End If
End Sub


