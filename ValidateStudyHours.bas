Attribute VB_Name = "ValidateStudyHours"
Sub ValidateStudyHours()
    ' Declare all necessary variables
    Dim targetWs As Worksheet
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim col As Long
    Dim row As Long
    Dim weekNumber As Integer
    Dim weekMap As Object ' Dictionary: column -> weekNumber
    Dim perSubjectHours As Object ' Dictionary: uniqueRecordKey -> weekNumber -> hours
    Dim perChildHours As Object ' Dictionary: uniqueChildKey -> weekNumber -> total hours
    Dim columnIHrs As Object ' Dictionary: childID -> column I hours
    Dim childID As Variant
    Dim subject As String
    Dim age As Integer
    Dim cellValue As Variant
    Dim dateInCell As Date
    Dim errorRows As Object ' Dictionary: row -> True
    Dim errorMessages As String
    Dim dictChildAge As Object ' Dictionary: childID -> age
    Dim headerDate As Variant
    Dim isFirstColumn As Boolean
    Dim errorChildKeys As Object ' Dictionary: uniqueChildKey -> True
    Dim headerColor As Long
    Dim valueB As Variant, valueC As Variant, valueE As Variant
    Dim uniqueRecordKey As String
    Dim uniqueChildKey As String
    Dim recordToRow As Object ' Dictionary: uniqueRecordKey -> row number
    Dim childToRecords As Object ' Dictionary: uniqueChildKey -> Collection of uniqueRecordKeys
    
    ' Initialize dictionaries using late binding
    Set weekMap = CreateObject("Scripting.Dictionary")
    Set perSubjectHours = CreateObject("Scripting.Dictionary")
    Set perChildHours = CreateObject("Scripting.Dictionary")
    Set columnIHrs = CreateObject("Scripting.Dictionary")
    Set errorRows = CreateObject("Scripting.Dictionary")
    Set dictChildAge = CreateObject("Scripting.Dictionary")
    Set errorChildKeys = CreateObject("Scripting.Dictionary")
    Set recordToRow = CreateObject("Scripting.Dictionary")
    Set childToRecords = CreateObject("Scripting.Dictionary")
    
    ' Improve performance by disabling screen updating and automatic calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Set worksheet to the active sheet
    Set targetWs = ActiveSheet ' Current sheet
    
    ' *** Step 1: Clear existing fill colors in columns A to H starting from row 11 ***
    With targetWs
        ' Determine the last row with data in column A
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        If lastRow < 11 Then
            MsgBox "No data found starting from row 11.", vbExclamation
            GoTo Cleanup
        End If
        
        ' Clear fill colors in columns A to H from row 11 to lastRow
        .Range("A11:H" & lastRow).Interior.ColorIndex = xlNone
    End With
    
    ' *** Step 2: Apply consistent color to columns D and E based on cell D10 ***
    With targetWs
        ' Get the color of cell D10
        headerColor = .Range("D10").Interior.Color
        ' Apply this color to columns D and E from row 11 to lastRow
        .Range("D11:E" & lastRow).Interior.Color = headerColor
    End With
    
    ' *** Step 3: Map columns J to AN to week numbers ***
    lastColumn = Columns("AN").Column ' Fixed at column 40
    weekNumber = 1
    isFirstColumn = True
    For col = Columns("J").Column To lastColumn
        headerDate = targetWs.Cells(5, col).value
        If IsDate(headerDate) Then
            dateInCell = CDate(headerDate)
            ' If the day is Monday and not the first column, increment week number
            If Weekday(dateInCell, vbMonday) = 1 Then ' vbMonday sets Monday as first day
                If Not isFirstColumn Then
                    weekNumber = weekNumber + 1
                End If
            End If
            weekMap(col) = weekNumber
            isFirstColumn = False
        Else
            ' If date is not valid, assign to current week
            weekMap(col) = weekNumber
        End If
    Next col
    
    ' *** Step 4: Loop through each row starting from 11 ***
    For row = 11 To lastRow
        ' Retrieve necessary cell values
        childID = Trim(targetWs.Cells(row, "A").value)
        subject = Trim(targetWs.Cells(row, "D").value)
        age = targetWs.Cells(row, "H").value
        
        ' Skip rows with empty childID or subject
        If childID = "" Or subject = "" Then
            GoTo NextRow
        End If
        
        ' Read additional columns B, C, E for unique record identification
        valueB = Trim(targetWs.Cells(row, "B").value)
        valueC = Trim(targetWs.Cells(row, "C").value)
        valueE = Trim(targetWs.Cells(row, "E").value)
        
        ' Create a unique key for the record by concatenating columns A, B, C, D, E
        uniqueRecordKey = childID & "|" & valueB & "|" & valueC & "|" & subject & "|" & valueE
        
        ' Map the uniqueRecordKey to the current row number
        If Not recordToRow.exists(uniqueRecordKey) Then
            recordToRow(uniqueRecordKey) = row
        End If
        
        ' Create a unique key for the child by concatenating columns A, B, C, H
        uniqueChildKey = childID & "|" & valueB & "|" & valueC & "|" & age
        
        ' Map the uniqueChildKey to its records
        If Not childToRecords.exists(uniqueChildKey) Then
            Set childToRecords(uniqueChildKey) = New Collection
        End If
        childToRecords(uniqueChildKey).Add uniqueRecordKey
        
        ' Store child age if not already stored
        If Not dictChildAge.exists(uniqueChildKey) Then
            dictChildAge(uniqueChildKey) = age
        End If
        
        ' Handle column I (previous month's hours for the child)
        Dim prevMonthHours As Variant
        prevMonthHours = targetWs.Cells(row, "I").value
        If IsNumeric(prevMonthHours) Then
            If Not columnIHrs.exists(uniqueChildKey) Then
                columnIHrs(uniqueChildKey) = prevMonthHours
            Else
                columnIHrs(uniqueChildKey) = columnIHrs(uniqueChildKey) + prevMonthHours
            End If
        End If
        
        ' Initialize perSubjectHours using uniqueRecordKey
        If Not perSubjectHours.exists(uniqueRecordKey) Then
            Set perSubjectHours(uniqueRecordKey) = CreateObject("Scripting.Dictionary")
        End If
        
        ' Initialize perChildHours using uniqueChildKey
        If Not perChildHours.exists(uniqueChildKey) Then
            Set perChildHours(uniqueChildKey) = CreateObject("Scripting.Dictionary")
        End If
        
        ' Loop through each column J to AN (10 to 40) to accumulate hours
        For col = Columns("J").Column To lastColumn
            Dim currentWeek As Integer
            currentWeek = weekMap(col)
            
            ' Initialize if not exist using .Add
            If Not perSubjectHours(uniqueRecordKey).exists(currentWeek) Then
                perSubjectHours(uniqueRecordKey).Add currentWeek, 0
            End If
            If Not perChildHours(uniqueChildKey).exists(currentWeek) Then
                perChildHours(uniqueChildKey).Add currentWeek, 0
            End If
            
            ' Get hours from the cell
            cellValue = targetWs.Cells(row, col).value
            If IsNumeric(cellValue) Then
                perSubjectHours(uniqueRecordKey)(currentWeek) = perSubjectHours(uniqueRecordKey)(currentWeek) + cellValue
                perChildHours(uniqueChildKey)(currentWeek) = perChildHours(uniqueChildKey)(currentWeek) + cellValue
            End If
        Next col
        
        ' Now, add column I's hours to week 1 per subject and per child
        If columnIHrs.exists(uniqueChildKey) Then
            ' Add to perSubjectHours for week 1
            If perSubjectHours(uniqueRecordKey).exists(1) Then
                perSubjectHours(uniqueRecordKey)(1) = perSubjectHours(uniqueRecordKey)(1) + columnIHrs(uniqueChildKey)
            Else
                perSubjectHours(uniqueRecordKey).Add 1, columnIHrs(uniqueChildKey)
            End If
            
            ' Add to perChildHours for week 1
            If perChildHours(uniqueChildKey).exists(1) Then
                perChildHours(uniqueChildKey)(1) = perChildHours(uniqueChildKey)(1) + columnIHrs(uniqueChildKey)
            Else
                perChildHours(uniqueChildKey).Add 1, columnIHrs(uniqueChildKey)
            End If
        End If
        
NextRow:
    Next row
    
    ' *** Step 5: Check per-subject per-week hours > 2 ***
    Dim uniqueRecordIter As Variant
    For Each uniqueRecordIter In perSubjectHours.keys
        Dim subjectWeeks As Object
        Set subjectWeeks = perSubjectHours(uniqueRecordIter)
        
        ' Extract childID and subject from uniqueRecordKey if needed
        Dim parts() As String
        parts = Split(uniqueRecordIter, "|")
        Dim currentChildID As String
        Dim currentSurname As String
        Dim currentName As String
        Dim currentSubject As String
        currentChildID = parts(0)
        currentSurname = parts(1)
        currentName = parts(2)
        currentSubject = parts(3)
        ' parts(4) is valueE, not needed here
        
        For Each wk In subjectWeeks.keys
            If Round(subjectWeeks(wk) / 45, 2) > 2 Then
                ' Retrieve the corresponding row number
                row = recordToRow(uniqueRecordIter)
                
                ' Record error for this row
                If Not errorRows.exists(row) Then
                    errorRows(row) = True
                End If
                
                ' Add to error message
                errorMessages = errorMessages & "Row " & row & ": Total hours for child " & currentChildID & " (" & currentSurname & " " & currentName & ") in subject '" & currentSubject & "' for week " & wk & " exceeds 2 hours." & vbCrLf
                
                ' No need to check further weeks for this record
                Exit For
            End If
        Next wk
    Next uniqueRecordIter
    
    ' *** Step 6: Check per-child per-week total hours ***
    Dim uniqueChildIter As Variant
    For Each uniqueChildIter In perChildHours.keys
        Dim childWeeks As Object
        Set childWeeks = perChildHours(uniqueChildIter)
        Dim childAge As Integer
        childAge = dictChildAge(uniqueChildIter)
        Dim limit As Integer
        If childAge <= 11 Then
            limit = 4
        Else
            limit = 6
        End If
        For Each wk In childWeeks.keys
            Dim totalHours As Double
            totalHours = Round((perChildHours(uniqueChildIter)(wk)) / 45, 2)
            If totalHours > limit Then
                ' Add to error message
                errorMessages = errorMessages & "Child " & Split(uniqueChildIter, "|")(0) & " (" & Split(uniqueChildIter, "|")(1) & " " & Split(uniqueChildIter, "|")(2) & "): Total hours for week " & wk & " (" & totalHours & ") exceeds the limit of " & limit & " hours." & vbCrLf
                
                ' Record uniqueChildKey for highlighting
                If Not errorChildKeys.exists(uniqueChildIter) Then
                    errorChildKeys(uniqueChildIter) = True
                End If
            End If
        Next wk
    Next uniqueChildIter
    
    ' *** Step 7: Highlight rows with per-subject errors by filling A:H with yellow ***
    If errorRows.Count > 0 Then
        Dim key As Variant
        For Each key In errorRows.keys
            targetWs.Range(targetWs.Cells(key, "A"), targetWs.Cells(key, "H")).Interior.Color = vbYellow
        Next key
    End If
    
    ' *** Step 8: Highlight all records for children with total hours exceeding limits by filling A:H with yellow ***
    If errorChildKeys.Count > 0 Then
        Dim recordsCollection As Collection
        Dim recKey As Variant
        For Each uniqueChildIter In errorChildKeys.keys
            ' Retrieve all uniqueRecordKeys associated with this uniqueChildKey
            Set recordsCollection = childToRecords(uniqueChildIter)
            For Each recKey In recordsCollection
                ' Retrieve the corresponding row number
                row = recordToRow(recKey)
                ' Highlight the row
                targetWs.Range(targetWs.Cells(row, "A"), targetWs.Cells(row, "H")).Interior.Color = vbYellow
            Next recKey
        Next uniqueChildIter
    End If
    
    ' *** Step 9: Show message box with errors ***
    If errorMessages <> "" Then
        MsgBox "Data validation completed with errors:" & vbCrLf & errorMessages, vbExclamation
    Else
        MsgBox "Data validation completed successfully. No errors found.", vbInformation
    End If
    
Cleanup:
    ' Restore Excel settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Cleanup objects
    Set weekMap = Nothing
    Set perSubjectHours = Nothing
    Set perChildHours = Nothing
    Set columnIHrs = Nothing
    Set errorRows = Nothing
    Set dictChildAge = Nothing
    Set errorChildKeys = Nothing
    Set recordToRow = Nothing
    Set childToRecords = Nothing
End Sub


