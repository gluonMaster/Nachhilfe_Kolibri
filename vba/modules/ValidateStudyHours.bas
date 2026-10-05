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
    Dim isCrossMonthWeek As Boolean
    Dim prevWs As Worksheet
    Dim prevLastWeekHours As Object ' Dictionary: crossKey(A|B|C|D|E) -> hours
    Dim prevChildCarryover As Object ' Dictionary: childKey(A|B|C) -> total carryover hours
    Dim crossMonthErrorRows As Object ' Dictionary: row -> True
    Dim prevCarryoverApplied As Object ' Dictionary: uniqueChildKey -> True
    Dim firstColDate As Variant
    Dim firstColIdx As Long
    Dim prevIdKeys As Object
    Dim prevChildKeys As Object
    Dim prevSubjectKeys As Object
    Dim prevFullKeys As Object
    Dim missingPrevEntries As Collection
    Dim missingPrevCount As Long
    Dim missingPrevLogSheetName As String
    Dim missingPrevLogWs As Worksheet

    ' Initialize dictionaries using late binding
    Set weekMap = CreateObject("Scripting.Dictionary")
    Set perSubjectHours = CreateObject("Scripting.Dictionary")
    Set perChildHours = CreateObject("Scripting.Dictionary")
    Set errorRows = CreateObject("Scripting.Dictionary")
    Set dictChildAge = CreateObject("Scripting.Dictionary")
    Set errorChildKeys = CreateObject("Scripting.Dictionary")
    Set recordToRow = CreateObject("Scripting.Dictionary")
    Set childToRecords = CreateObject("Scripting.Dictionary")
    Set prevLastWeekHours = CreateObject("Scripting.Dictionary")
    Set prevChildCarryover = CreateObject("Scripting.Dictionary")
    Set crossMonthErrorRows = CreateObject("Scripting.Dictionary")
    Set prevCarryoverApplied = CreateObject("Scripting.Dictionary")
    Set missingPrevEntries = New Collection
    missingPrevCount = 0
    missingPrevLogSheetName = ""
    
    ' Improve performance by disabling screen updating and automatic calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Set worksheet to the active sheet
    Set targetWs = ActiveSheet ' Current sheet
    
    ' *** Step 1: Clear existing fill colors in columns A to H starting from row 11 ***
    With targetWs
        ' Determine the last row with data in column A
        lastRow = .Cells(.rows.Count, "A").End(xlUp).row
        If lastRow < 11 Then
            MsgBox "No data found starting from row 11.", vbExclamation
            GoTo CleanUp
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

    ' Determine whether week 1 is a cross-month week
    isCrossMonthWeek = False
    firstColDate = Empty
    For firstColIdx = Columns("J").Column To lastColumn
        firstColDate = targetWs.Cells(5, firstColIdx).value
        If IsDate(firstColDate) Then Exit For
    Next firstColIdx

    If IsDate(firstColDate) Then
        Dim firstDayOfWeek As Integer
        firstDayOfWeek = Weekday(CDate(firstColDate), vbMonday)
        ' Cross-month week: month starts Tue–Sat (2–6).
        ' Monday (1) = week starts fresh; Sunday (7) = no study load, skip.
        If firstDayOfWeek >= 2 And firstDayOfWeek <= 6 Then
            isCrossMonthWeek = True
        End If
    End If

    ' If week 1 continues from previous month, ask user whether previous month data is present
    If isCrossMonthWeek Then
        Dim prevMonthAnswer As Integer
        prevMonthAnswer = MsgBox( _
            "A cross-month week was detected: the current month does not start on Monday." & vbCrLf & _
            "Does this workbook contain data for the previous month?", _
            vbQuestion + vbYesNo + vbDefaultButton1, "Previous Month Data")

        If prevMonthAnswer = vbYes Then
            Set prevWs = GetPrevMonthSheet(targetWs)
            If prevWs Is Nothing Then
                Dim expectedPrevFirstText As String
                If IsDate(targetWs.Range("A1").value) Then
                    expectedPrevFirstText = Format$(DateSerial(Year(CDate(targetWs.Range("A1").value)), Month(CDate(targetWs.Range("A1").value)) - 1, 1), "yyyy-mm-dd")
                Else
                    expectedPrevFirstText = "N/A (A1 is not a valid date)"
                End If
                MsgBox "Validation aborted: previous month sheet was not found." & vbCrLf & _
                       "Current sheet: " & targetWs.name & vbCrLf & _
                       "Expected previous month first date in A1: " & expectedPrevFirstText, vbCritical
                GoTo CleanUp
            End If

            VSH_BuildPrevKeyIndex prevWs, prevIdKeys, prevChildKeys, prevSubjectKeys, prevFullKeys
            Set prevLastWeekHours = BuildPrevLastWeekHours(prevWs)

            ' Aggregate carryover per child (A|B|C) to avoid per-child undercounting.
            Dim prevCrossKey As Variant
            For Each prevCrossKey In prevLastWeekHours.Keys
                Dim prevParts() As String
                prevParts = Split(CStr(prevCrossKey), "|")
                If UBound(prevParts) >= 2 Then
                    Dim prevChildKey As String
                    prevChildKey = prevParts(0) & "|" & prevParts(1) & "|" & prevParts(2)
                    If prevChildCarryover.exists(prevChildKey) Then
                        prevChildCarryover(prevChildKey) = prevChildCarryover(prevChildKey) + CDbl(prevLastWeekHours(prevCrossKey))
                    Else
                        prevChildCarryover.Add prevChildKey, CDbl(prevLastWeekHours(prevCrossKey))
                    End If
                End If
            Next prevCrossKey
        Else
            ' User confirmed no previous month data - skip all cross-month processing
            isCrossMonthWeek = False
        End If
    End If

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

        ' Cross-month diagnostics: collect records that do not match previous month by full key.
        If isCrossMonthWeek Then
            If Not prevFullKeys Is Nothing Then
                If Not prevFullKeys.exists(uniqueRecordKey) Then
                    Dim idKey As String
                    Dim childKey As String
                    Dim subjectKey As String
                    Dim missingReason As String

                    idKey = CStr(childID)
                    childKey = idKey & "|" & CStr(valueB) & "|" & CStr(valueC)
                    subjectKey = childKey & "|" & CStr(subject)

                    missingReason = VSH_GetMissingMatchReason( _
                        idKey, _
                        childKey, _
                        subjectKey, _
                        uniqueRecordKey, _
                        prevIdKeys, _
                        prevChildKeys, _
                        prevSubjectKeys, _
                        prevFullKeys)

                    VSH_AddMissingPrevEntry missingPrevEntries, _
                        row, _
                        CStr(childID), _
                        CStr(valueB), _
                        CStr(valueC), _
                        CStr(subject), _
                        CStr(valueE), _
                        missingReason
                End If
            End If
        End If

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

        ' Cross-month carryover: add previous month last-week hours to week 1
        ' uniqueRecordKey (A|B|C|D|E) is the same key used in BuildPrevLastWeekHours
        If isCrossMonthWeek And prevLastWeekHours.exists(uniqueRecordKey) Then
            Dim prevHrs As Double
            prevHrs = CDbl(prevLastWeekHours(uniqueRecordKey))

            ' Add carryover to per-subject week 1
            If perSubjectHours(uniqueRecordKey).exists(1) Then
                perSubjectHours(uniqueRecordKey)(1) = perSubjectHours(uniqueRecordKey)(1) + prevHrs
            Else
                perSubjectHours(uniqueRecordKey).Add 1, prevHrs
            End If
        End If

        If isCrossMonthWeek Then
            ' Add carryover to per-child week 1 only once per child
            If Not prevCarryoverApplied.exists(uniqueChildKey) Then
                Dim childCarryKey As String
                Dim prevChildHrs As Double
                childCarryKey = childID & "|" & valueB & "|" & valueC
                prevChildHrs = 0
                If prevChildCarryover.exists(childCarryKey) Then
                    prevChildHrs = CDbl(prevChildCarryover(childCarryKey))
                End If

                If prevChildHrs > 0 Then
                    If perChildHours(uniqueChildKey).exists(1) Then
                        perChildHours(uniqueChildKey)(1) = perChildHours(uniqueChildKey)(1) + prevChildHrs
                    Else
                        perChildHours(uniqueChildKey).Add 1, prevChildHrs
                    End If
                End If
                prevCarryoverApplied(uniqueChildKey) = True
            End If
        End If

NextRow:
    Next row
    
    ' *** Step 5: Check per-subject per-week hours > 2 ***
    Dim uniqueRecordIter As Variant
    For Each uniqueRecordIter In perSubjectHours.Keys
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
        
        For Each wk In subjectWeeks.Keys
            If Round(subjectWeeks(wk) / 45, 2) > 2 Then
                ' Retrieve the corresponding row number
                row = recordToRow(uniqueRecordIter)
                
                ' Record error for this row
                If Not errorRows.exists(row) Then
                    errorRows(row) = True
                End If

                Dim subjectMsgPrefix As String
                subjectMsgPrefix = ""
                If CLng(wk) = 1 And isCrossMonthWeek Then
                    ' Only mark as cross-month if there were actual carryover hours for this record
                    If prevLastWeekHours.exists(uniqueRecordIter) Then
                        If Not crossMonthErrorRows.exists(row) Then
                            crossMonthErrorRows(row) = True
                        End If
                        subjectMsgPrefix = "[CROSS-MONTH] "
                    End If
                End If
                
                ' Add to error message
                errorMessages = errorMessages & subjectMsgPrefix & "Row " & row & ": Total hours for child " & currentChildID & " (" & currentSurname & " " & currentName & ") in subject '" & currentSubject & "' for week " & wk & " exceeds 2 hours." & vbCrLf
                
                ' No need to check further weeks for this record
                Exit For
            End If
        Next wk
    Next uniqueRecordIter
    
    ' *** Step 6: Check per-child per-week total hours ***
    Dim uniqueChildIter As Variant
    For Each uniqueChildIter In perChildHours.Keys
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
        For Each wk In childWeeks.Keys
            Dim totalHours As Double
            totalHours = Round((perChildHours(uniqueChildIter)(wk)) / 45, 2)
            If totalHours > limit Then
                Dim childMsgPrefix As String
                childMsgPrefix = ""
                If CLng(wk) = 1 And isCrossMonthWeek Then
                    ' Only mark as cross-month if there were actual carryover hours for this child
                    Dim ccParts() As String
                    ccParts = Split(CStr(uniqueChildIter), "|")
                    Dim ccKey As String
                    ccKey = ccParts(0) & "|" & ccParts(1) & "|" & ccParts(2)
                    If prevChildCarryover.exists(ccKey) Then
                        childMsgPrefix = "[CROSS-MONTH] "

                        ' Mark all rows of this child as cross-month errors
                        Dim crossRecords As Collection
                        Set crossRecords = childToRecords(uniqueChildIter)
                        Dim crossRec As Variant
                        For Each crossRec In crossRecords
                            Dim crossRow As Long
                            crossRow = recordToRow(crossRec)
                            If Not crossMonthErrorRows.exists(crossRow) Then
                                crossMonthErrorRows(crossRow) = True
                            End If
                        Next crossRec
                    End If
                End If

                ' Add to error message
                errorMessages = errorMessages & childMsgPrefix & "Child " & Split(uniqueChildIter, "|")(0) & " (" & Split(uniqueChildIter, "|")(1) & " " & Split(uniqueChildIter, "|")(2) & "): Total hours for week " & wk & " (" & totalHours & ") exceeds the limit of " & limit & " hours." & vbCrLf
                
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
        For Each key In errorRows.Keys
            If Not crossMonthErrorRows.exists(key) Then
                targetWs.Range(targetWs.Cells(key, "A"), targetWs.Cells(key, "H")).Interior.Color = vbYellow
            End If
        Next key
    End If

    ' *** Step 7b: Highlight cross-month errors with amber color ***
    If crossMonthErrorRows.Count > 0 Then
        For Each key In crossMonthErrorRows.Keys
            targetWs.Range(targetWs.Cells(key, "A"), targetWs.Cells(key, "H")).Interior.Color = RGB(255, 180, 0)
        Next key
    End If

    ' *** Step 8: Highlight all records for children with total hours exceeding limits by filling A:H with yellow ***
    If errorChildKeys.Count > 0 Then
        Dim recordsCollection As Collection
        Dim recKey As Variant
        For Each uniqueChildIter In errorChildKeys.Keys
            ' Retrieve all uniqueRecordKeys associated with this uniqueChildKey
            Set recordsCollection = childToRecords(uniqueChildIter)
            For Each recKey In recordsCollection
                ' Retrieve the corresponding row number
                row = recordToRow(recKey)
                ' Highlight the row with cross-month priority
                If crossMonthErrorRows.exists(row) Then
                    targetWs.Range(targetWs.Cells(row, "A"), targetWs.Cells(row, "H")).Interior.Color = RGB(255, 180, 0)
                Else
                    targetWs.Range(targetWs.Cells(row, "A"), targetWs.Cells(row, "H")).Interior.Color = vbYellow
                End If
            Next recKey
        Next uniqueChildIter
    End If

    ' Cross-month diagnostics output: write potential missing matches to a temporary log sheet.
    If isCrossMonthWeek Then
        missingPrevCount = missingPrevEntries.Count
        If missingPrevCount > 0 Then
            Set missingPrevLogWs = VSH_CreateMissingPrevLogSheet(targetWs)
            If Not missingPrevLogWs Is Nothing Then
                VSH_WriteMissingPrevEntries missingPrevLogWs, targetWs, prevWs, missingPrevEntries
                missingPrevLogSheetName = missingPrevLogWs.name
            End If
        End If
    End If
    
    ' *** Step 9: Show message box with errors ***
    Dim missingLogNote As String
    missingLogNote = ""
    If missingPrevCount > 0 Then
        If missingPrevLogSheetName <> "" Then
            missingLogNote = vbCrLf & "Potential previous-month mismatches were logged to sheet '" & missingPrevLogSheetName & "'."
        Else
            missingLogNote = vbCrLf & "Potential previous-month mismatches were detected, but log sheet creation failed."
        End If
    End If

    If errorMessages <> "" Then
        MsgBox "Data validation completed with errors:" & vbCrLf & errorMessages & missingLogNote, vbExclamation
    Else
        If missingPrevCount > 0 Then
            MsgBox "Data validation completed. No limit violations found." & missingLogNote, vbExclamation
        Else
            MsgBox "Data validation completed successfully. No errors found.", vbInformation
        End If
    End If
    
CleanUp:
    ' Restore Excel settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Cleanup objects
    Set weekMap = Nothing
    Set perSubjectHours = Nothing
    Set perChildHours = Nothing
    Set prevLastWeekHours = Nothing
    Set prevChildCarryover = Nothing
    Set crossMonthErrorRows = Nothing
    Set prevCarryoverApplied = Nothing
    Set errorRows = Nothing
    Set dictChildAge = Nothing
    Set errorChildKeys = Nothing
    Set recordToRow = Nothing
    Set childToRecords = Nothing
    Set prevIdKeys = Nothing
    Set prevChildKeys = Nothing
    Set prevSubjectKeys = Nothing
    Set prevFullKeys = Nothing
    Set missingPrevEntries = Nothing
    If Not missingPrevLogWs Is Nothing Then Set missingPrevLogWs = Nothing
    If Not prevWs Is Nothing Then Set prevWs = Nothing
End Sub

Private Function GetPrevMonthSheet(currentWs As Worksheet) As Worksheet
    ' Find worksheet where A1 equals first day of the previous month of currentWs.A1
    Dim refDate As Date
    If Not IsDate(currentWs.Range("A1").value) Then Exit Function
    refDate = CDate(currentWs.Range("A1").value)
    If Day(refDate) <> 1 Then Exit Function

    Dim prevFirst As Date
    prevFirst = DateSerial(Year(refDate), Month(refDate) - 1, 1)

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsDate(ws.Range("A1").value) Then
            If CDate(ws.Range("A1").value) = prevFirst Then
                Set GetPrevMonthSheet = ws
                Exit Function
            End If
        End If
    Next ws
End Function

Private Function BuildPrevLastWeekHours(prevWs As Worksheet) As Object
    ' Returns dictionary: crossKey(A|B|C|D|E) -> total hours in the last week of prevWs
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim prevLastCol As Long
    prevLastCol = prevWs.Columns("AN").Column

    ' Build week map for previous month sheet
    Dim prevWeekMap As Object
    Set prevWeekMap = CreateObject("Scripting.Dictionary")
    Dim wn As Integer
    Dim isFirst As Boolean
    Dim c As Long

    wn = 1
    isFirst = True

    For c = prevWs.Columns("J").Column To prevLastCol
        Dim hd As Variant
        hd = prevWs.Cells(5, c).value
        If IsDate(hd) Then
            If Weekday(CDate(hd), vbMonday) = 1 Then
                If Not isFirst Then wn = wn + 1
            End If
            prevWeekMap(c) = wn
            isFirst = False
        Else
            prevWeekMap(c) = wn
        End If
    Next c

    Dim maxWeek As Integer
    maxWeek = wn

    ' Sum hours of the last week per record
    Dim prevLastRow As Long
    prevLastRow = prevWs.Cells(prevWs.rows.Count, "A").End(xlUp).row
    If prevLastRow < 11 Then
        Set BuildPrevLastWeekHours = result
        Exit Function
    End If

    Dim r As Long
    For r = 11 To prevLastRow
        Dim cID As String, cB As String, cC As String, cD As String, cE As String
        cID = Trim(prevWs.Cells(r, "A").value)
        cB = Trim(prevWs.Cells(r, "B").value)
        cC = Trim(prevWs.Cells(r, "C").value)
        cD = Trim(prevWs.Cells(r, "D").value)
        cE = Trim(prevWs.Cells(r, "E").value)
        If cID = "" Or cD = "" Then GoTo NextPrevRow

        Dim crossKey As String
        crossKey = cID & "|" & cB & "|" & cC & "|" & cD & "|" & cE

        Dim hrs As Double
        hrs = 0
        For c = prevWs.Columns("J").Column To prevLastCol
            If prevWeekMap.exists(c) Then
                If prevWeekMap(c) = maxWeek Then
                    Dim cv As Variant
                    cv = prevWs.Cells(r, c).value
                    If IsNumeric(cv) Then
                        hrs = hrs + CDbl(cv)
                    End If
                End If
            End If
        Next c

        If hrs > 0 Then
            If result.exists(crossKey) Then
                result(crossKey) = result(crossKey) + hrs
            Else
                result.Add crossKey, hrs
            End If
        End If

NextPrevRow:
    Next r

    Set BuildPrevLastWeekHours = result
End Function

