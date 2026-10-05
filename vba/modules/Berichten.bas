Attribute VB_Name = "Berichten"
Option Explicit

Sub GenerateChildReportsWithDetailedTables()
    ' Define variables
    Dim wsLoad As Worksheet
    Dim wsTemplate As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim reportFolder As String
    Dim childKey As Variant
    Dim dictChildren As Object
    Dim childData As Variant
    Dim i As Long
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim baseFileName As String
    Dim fileNamePDF As String
    Dim fileNameXLSX As String
    Dim childLastName As String
    Dim childFirstName As String
    Dim childBirthDate As Variant
    Dim socialServiceID As String
    Dim lessonStartDate As Variant
    Dim lessonEndDate As Variant
    Dim templateRange As Range
    Dim pdfPath As String
    Dim excelPath As String
    Dim fso As Object
    Dim recordCollection As Collection
    Dim recordData As Variant
    Dim templateHeaderRange As Range
    Dim templateRowRange As Range
    Dim templateFooterRange As Range
    Dim lineNumber As Long
    Dim disciplineName As String
    Dim lessonTypeCode As String
    Dim lessonTypeString As String
    Dim studyHourValue As Variant
    Dim dateValue As Variant
    Dim calculatedValueC As Long
    Dim calculatedValueF As Double
    Dim costPerHour As Double
    Dim totalCostFromRecord As Double
    Dim totalCostAllDisciplines As Double
    Dim totalHoursFromRecord As Double
    Dim totalChildren As Long
    Dim processedChildren As Long
    Dim reportDate As Date
    Dim monthNumber As Integer
    Dim yearNumber As Integer
    Dim firstLetter As String
    Dim subfolderPath As String
    Dim processSelected As Boolean
    Dim rowsToProcess As Range
    Dim selectedRow As Range
    Dim isValidSelection As Boolean
    
    ' Initialize
    Set wsLoad = ThisWorkbook.ActiveSheet ' Assumes the user is on the monthly load sheet
    Set wsTemplate = ThisWorkbook.Sheets("Shablon") ' Template sheet
    Set dictChildren = CreateObject("Scripting.Dictionary") ' Late binding
    
    ' Determine the last row with data in column A (Child ID)
    lastRow = wsLoad.Cells(wsLoad.rows.Count, "A").End(xlUp).row
    If lastRow < 11 Then
        MsgBox "No data found starting from row 11.", vbExclamation
        Exit Sub
    End If
    
    ' Ask the user if they want to process only selected rows
    If MsgBox("Do you want to generate reports only for the selected records?", vbYesNo + vbQuestion, "Generate Reports") = vbYes Then
        processSelected = True
        ' Check if there are selected rows
        If TypeName(selection) <> "Range" Then
            MsgBox "Please select the rows for which you want to generate reports.", vbExclamation
            processSelected = False
        Else
            ' Check that only entire rows are selected and they start from row 11
            Set rowsToProcess = Nothing
            For Each selectedRow In selection.rows
                If selectedRow.row < 11 Or selectedRow.row > lastRow Then
                    MsgBox "Selected rows are outside the data range (starting from row 11).", vbExclamation
                    processSelected = False
                    Exit For
                Else
                    If rowsToProcess Is Nothing Then
                        Set rowsToProcess = selectedRow
                    Else
                        Set rowsToProcess = Union(rowsToProcess, selectedRow)
                    End If
                End If
            Next selectedRow
            If rowsToProcess Is Nothing Then
                MsgBox "No valid selected rows to process.", vbExclamation
                processSelected = False
            End If
        End If
    Else
        processSelected = False
    End If
    
    ' Prompt user to select the destination folder
    With Application.fileDialog(msoFileDialogFolderPicker)
        .title = "Select Destination Folder for Reports"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "No folder selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
        reportFolder = .SelectedItems(1)
    End With
    
    ' Initialize FileSystemObject for handling file paths
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Disable Screen Updating and other settings to prevent flickering
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' Build a dictionary of children
    ' Key: Concatenation of columns A, B, C, F, G, H
    ' Value: Collection of records (each record is an array with necessary data)
    
    If processSelected Then
        ' Process only selected rows
        For Each selectedRow In rowsToProcess.rows
            currentRow = selectedRow.row
            ' Read necessary cells
            Dim cellA As String
            Dim cellB As String
            Dim cellC As String
            Dim cellD As String
            Dim cellE As String
            Dim cellF As Variant
            Dim cellG As Variant
            Dim cellH As Variant
            Dim cellAU As String
            Dim cellAV As Variant
            Dim cellAP As Double
            Dim cellAQ As Double
            
            cellA = Trim(wsLoad.Cells(currentRow, "A").value) ' Child ID
            cellB = Trim(wsLoad.Cells(currentRow, "B").value) ' Last Name
            cellC = Trim(wsLoad.Cells(currentRow, "C").value) ' First Name
            cellD = Trim(wsLoad.Cells(currentRow, "D").value) ' Discipline
            cellE = Trim(wsLoad.Cells(currentRow, "E").value) ' Lesson Type
            cellF = wsLoad.Cells(currentRow, "F").value ' Start Date
            cellG = wsLoad.Cells(currentRow, "G").value ' End Date
            cellH = wsLoad.Cells(currentRow, "H").value ' Age
            cellAU = Trim(wsLoad.Cells(currentRow, "AU").value) ' Social Service ID
            cellAV = wsLoad.Cells(currentRow, "AV").value ' Birth Date
            cellAP = wsLoad.Cells(currentRow, "AP").value ' Cost per Hour
            cellAQ = wsLoad.Cells(currentRow, "AQ").value ' Total Cost
            
            ' Skip rows with missing critical data
            If cellA = "" Or cellB = "" Or cellC = "" Or cellD = "" Or cellE = "" Or cellF = "" Or cellG = "" Or cellH = "" Then
                ' Log skipped rows
                Call LogSkippedRow(currentRow)
                GoTo NextSelectedRow
            End If
            
            ' Create a unique key for each child
            childKey = cellA & "|" & cellB & "|" & cellC
            
            ' If the child is not yet in the dictionary, add them with a new collection
            If Not dictChildren.exists(childKey) Then
                Set recordCollection = New Collection
                dictChildren.Add childKey, recordCollection
            Else
                Set recordCollection = dictChildren(childKey)
            End If
            
            ' Add the current record to the child's collection
            ' Record data includes: Discipline, Lesson Type, Cost per Hour, Total Cost, Social Service ID, Birth Date, Row Number
            recordCollection.Add Array(cellD, cellE, cellAP, cellAQ, cellAU, cellAV, currentRow, cellF, cellG)
        
NextSelectedRow:
        Next selectedRow
    Else
        ' Process all rows starting from row 11
        For currentRow = 11 To lastRow
            ' Read necessary cells
            
            cellA = Trim(wsLoad.Cells(currentRow, "A").value) ' Child ID
            cellB = Trim(wsLoad.Cells(currentRow, "B").value) ' Last Name
            cellC = Trim(wsLoad.Cells(currentRow, "C").value) ' First Name
            cellD = Trim(wsLoad.Cells(currentRow, "D").value) ' Discipline
            cellE = Trim(wsLoad.Cells(currentRow, "E").value) ' Lesson Type
            cellF = wsLoad.Cells(currentRow, "F").value ' Start Date
            cellG = wsLoad.Cells(currentRow, "G").value ' End Date
            cellH = wsLoad.Cells(currentRow, "H").value ' Age
            cellAU = Trim(wsLoad.Cells(currentRow, "AU").value) ' Social Service ID
            cellAV = wsLoad.Cells(currentRow, "AV").value ' Birth Date
            cellAP = wsLoad.Cells(currentRow, "AP").value ' Cost per Hour
            cellAQ = wsLoad.Cells(currentRow, "AQ").value ' Total Cost
            
            ' Skip rows with missing critical data
            If cellA = "" Or cellB = "" Or cellC = "" Or cellD = "" Or cellE = "" Or cellF = "" Or cellG = "" Or cellH = "" Then
                ' Log skipped rows
                Call LogSkippedRow(currentRow)
                GoTo NextRow
            End If
            
            ' Create a unique key for each child
            childKey = cellA & "|" & cellB & "|" & cellC
            
            ' If the child is not yet in the dictionary, add them with a new collection
            If Not dictChildren.exists(childKey) Then
                Set recordCollection = New Collection
                dictChildren.Add childKey, recordCollection
            Else
                Set recordCollection = dictChildren(childKey)
            End If
            
            ' Add the current record to the child's collection
            ' Record data includes: Discipline, Lesson Type, Cost per Hour, Total Cost, Social Service ID, Birth Date, Row Number
            recordCollection.Add Array(cellD, cellE, cellAP, cellAQ, cellAU, cellAV, currentRow, cellF, cellG)
        
NextRow:
        Next currentRow
    End If
    
    ' General settings for progress bar
    totalChildren = dictChildren.Count
    processedChildren = 0
    
    ' Initialize and show the progress form
    With frmProgress
        .lblProgress.Caption = ""
        .lblStatus.Caption = "Starting report generation..."
        .fraProgress.Width = 433 ' Ensure this matches the design
        .lblProgress.Width = 0
        .cancelRequested = False
        .Show vbModeless
    End With
    
    ' Iterate through each child and generate reports
    For Each childKey In dictChildren.Keys
        ' Check if cancellation was requested
        If frmProgress.cancelRequested Then
            MsgBox "Operation cancelled by the user.", vbInformation, "Cancelled"
            Exit For
        End If
        
        ' Increment processed children count
        processedChildren = processedChildren + 1
        
        ' Calculate progress percentage
        Dim progressPercent As Integer
        progressPercent = Int((processedChildren / totalChildren) * 100)
        If progressPercent > 100 Then progressPercent = 100
        
        ' Update progress bar
        UpdateProgressBar progressPercent
        frmProgress.lblStatus.Caption = "Processing " & processedChildren & " of " & totalChildren & " children..."
        DoEvents ' Allow the form to update
        
        ' Retrieve child data from key
        Dim splitKey() As String
        splitKey = Split(childKey, "|")
        childLastName = splitKey(1) ' Last Name
        childFirstName = splitKey(2) ' First Name
        
        ' Initialize total cost for all disciplines
        totalCostAllDisciplines = 0
        
        ' Determine the first letter of the last name
        firstLetter = Left(childLastName, 1)
        firstLetter = UCase(firstLetter)
        ' Check if firstLetter is a letter A-Z
        If firstLetter < "A" Or firstLetter > "Z" Then
            firstLetter = "Others"
        End If
        
        ' Define subfolder path
        subfolderPath = fso.BuildPath(reportFolder, firstLetter)
        ' Create subfolder if it doesn't exist
        If Not fso.FolderExists(subfolderPath) Then
            fso.CreateFolder subfolderPath
        End If
        
        ' Create a new workbook and hide it
        Set wbNew = Workbooks.Add(xlWBATWorksheet) ' Create a new workbook with one sheet
        wbNew.Windows(1).Visible = False ' Hide the new workbook
        Set wsNew = wbNew.Sheets(1)
        
        ' Copy the initial template range (A1:F9) from Template to the new workbook
        Set templateRange = wsTemplate.Range("A1:F9")
        templateRange.Copy
        wsNew.Range("A1").PasteSpecial Paste:=xlPasteAll
        
        ' Copy the column widths from Template to the new worksheet
        wsTemplate.Columns("A:F").Copy
        wsNew.Columns("A:F").PasteSpecial Paste:=xlPasteColumnWidths
        
        ' Populate specific cells with child data
        Dim combinedName As String
        combinedName = childLastName & ", " & childFirstName
        wsNew.Range("C2").value = combinedName
        wsNew.Range("E3").value = dictChildren(childKey)(1)(5) ' Birth Date from first record
        wsNew.Range("C3").value = dictChildren(childKey)(1)(4) ' Social Service ID from first record
        ' Collect unique authorization periods from all records of this child
        Dim allPeriods(1, 1) As Variant
        Dim periodCount As Integer
        Dim recIdx As Long
        Dim pIdx As Integer
        Dim periodFound As Boolean
        periodCount = 0

        For recIdx = 1 To dictChildren(childKey).Count
            Dim recF As Variant
            Dim recG As Variant
            recF = dictChildren(childKey)(recIdx)(7)
            recG = dictChildren(childKey)(recIdx)(8)
            periodFound = False
            For pIdx = 0 To periodCount - 1
                If allPeriods(pIdx, 0) = recF And allPeriods(pIdx, 1) = recG Then
                    periodFound = True
                    Exit For
                End If
            Next pIdx
            If Not periodFound And periodCount < 2 Then
                allPeriods(periodCount, 0) = recF
                allPeriods(periodCount, 1) = recG
                periodCount = periodCount + 1
            End If
        Next recIdx

        ' Sort periods by start date ascending, then by end date ascending
        If periodCount = 2 Then
            If allPeriods(0, 0) > allPeriods(1, 0) Or _
               (allPeriods(0, 0) = allPeriods(1, 0) And allPeriods(0, 1) > allPeriods(1, 1)) Then
                Dim swapVal As Variant
                swapVal = allPeriods(0, 0)
                allPeriods(0, 0) = allPeriods(1, 0)
                allPeriods(1, 0) = swapVal
                swapVal = allPeriods(0, 1)
                allPeriods(0, 1) = allPeriods(1, 1)
                allPeriods(1, 1) = swapVal
            End If
        End If

        ' Populate header cells based on number of unique periods
        If periodCount = 1 Then
            wsNew.Range("C7").value = allPeriods(0, 0)
            wsNew.Range("C8").value = allPeriods(0, 1)
        ElseIf periodCount = 2 Then
            wsNew.Range("C6").value = "von " & Format(CDate(allPeriods(0, 0)), "dd.mm.yyyy") & _
                                      " bis " & Format(CDate(allPeriods(0, 1)), "dd.mm.yyyy") & " und"
            wsNew.Range("C7").value = allPeriods(1, 0)
            wsNew.Range("C8").value = allPeriods(1, 1)
        End If
        
        ' -------------------------------------------------------
        ' Retrieve the child's ID from splitKey(0)
        Dim wsKinder As Worksheet
        Dim foundCell As Range
        Dim foundRow As Long
        Dim childID As String
        
        Set wsKinder = ThisWorkbook.Sheets("Kinder")
        
        childID = splitKey(0) ' child's ID in format "X. XXXX"
        
        ' Find the child's address on the Kinder sheet
        Set foundCell = wsKinder.Columns("B").Find(What:=childID, LookAt:=xlWhole, MatchCase:=False)
        If Not foundCell Is Nothing Then
            foundRow = foundCell.row
            ' Place the address parts into E4 and E5
            wsNew.Range("E4").value = wsKinder.Cells(foundRow, 19).value
            wsNew.Range("E5").value = wsKinder.Cells(foundRow, 20).value
            
            ' Format the cells E4 and E5
            With wsNew.Range("E4:E5")
                .HorizontalAlignment = xlLeft
                .Font.Bold = True
                .Font.name = "Calibri"
                .Font.Size = 10
            End With
        Else
            ' If no match is found, you can leave these cells blank or handle it differently if needed
            wsNew.Range("E4").value = ""
            wsNew.Range("E5").value = ""
        End If
        
        ' -------------------------------------------------------
        
        ' Format the date cells as desired (e.g., dd.mm.yyyy)
        wsNew.Range("E3").NumberFormat = "dd.mm.yyyy"
        wsNew.Range("C7").NumberFormat = "dd.mm.yyyy"
        wsNew.Range("C8").NumberFormat = "dd.mm.yyyy"
        
        ' Initialize lineNumber for table entries
        ' Assuming that after A1:F12, the tables start from row 10
        lineNumber = 10
        
        ' Group all records of this child by discipline and lesson type
        Dim dictDisciplines As Object
        Dim discRecords As Collection
        Dim discGroupKey As Variant
        Dim discKey As String
        Set dictDisciplines = CreateObject("Scripting.Dictionary")

        For i = 1 To dictChildren(childKey).Count
            recordData = dictChildren(childKey)(i)
            discKey = recordData(0) & "|" & recordData(1)
            If Not dictDisciplines.exists(discKey) Then
                Set discRecords = New Collection
                dictDisciplines.Add discKey, discRecords
            End If
            dictDisciplines(discKey).Add dictChildren(childKey)(i)
        Next i

        ' Iterate through each discipline group
        For Each discGroupKey In dictDisciplines.Keys
            Dim discCollection As Collection
            Set discCollection = dictDisciplines(discGroupKey)

            ' Get discipline info from the first record of this group
            recordData = discCollection(1)
            disciplineName = recordData(0) ' Discipline
            lessonTypeCode = recordData(1) ' Lesson Type (G or I)
            costPerHour = recordData(2)    ' Cost per Hour

            ' Determine lesson type string
            If lessonTypeCode = "G" Then
                lessonTypeString = "Gruppenunterricht"
            ElseIf lessonTypeCode = "I" Then
                lessonTypeString = "Einzelunterricht"
            Else
                lessonTypeString = "Unknown Type"
            End If

            ' Create the header string "Discipline Name / Lesson Type"
            Dim headerString As String
            headerString = disciplineName & " / " & lessonTypeString

            ' Copy the table header from Template sheet (B10:F11)
            Set templateHeaderRange = wsTemplate.Range("B10:F11")
            templateHeaderRange.Copy
            wsNew.Range("B" & lineNumber).PasteSpecial Paste:=xlPasteAll

            ' Populate the header string
            wsNew.Range("C" & lineNumber).value = headerString

            ' Move to the next line for table rows (header occupies 2 rows)
            lineNumber = lineNumber + 2

            ' Collect all non-zero entries from all records of this discipline group
            Dim col As Long
            Dim discRecIdx As Long
            Dim entryCount As Long
            entryCount = 0

            ' First pass: count entries
            For discRecIdx = 1 To discCollection.Count
                recordData = discCollection(discRecIdx)
                currentRow = recordData(6)
                For col = 10 To 40
                    studyHourValue = Round(wsLoad.Cells(currentRow, col).value / 45, 2)
                    If IsNumeric(studyHourValue) And studyHourValue > 0 Then
                        entryCount = entryCount + 1
                    End If
                Next col
            Next discRecIdx

            totalHoursFromRecord = 0
            totalCostFromRecord = 0

            If entryCount > 0 Then
                ' Allocate arrays for collected entries
                Dim entryDates() As Variant
                Dim entryHours() As Double
                ReDim entryDates(1 To entryCount)
                ReDim entryHours(1 To entryCount)

                ' Second pass: fill arrays
                Dim entryIdx As Long
                entryIdx = 0
                For discRecIdx = 1 To discCollection.Count
                    recordData = discCollection(discRecIdx)
                    currentRow = recordData(6)
                    For col = 10 To 40
                        studyHourValue = Round(wsLoad.Cells(currentRow, col).value / 45, 2)
                        If IsNumeric(studyHourValue) And studyHourValue > 0 Then
                            entryIdx = entryIdx + 1
                            entryDates(entryIdx) = wsLoad.Cells(5, col).value
                            entryHours(entryIdx) = studyHourValue
                        End If
                    Next col
                Next discRecIdx

                ' Sort entries by date ascending (bubble sort)
                Dim sortM As Long, sortN As Long
                Dim tmpDate As Variant, tmpHour As Double
                For sortM = 1 To entryCount - 1
                    For sortN = sortM + 1 To entryCount
                        If entryDates(sortN) < entryDates(sortM) Then
                            tmpDate = entryDates(sortM)
                            entryDates(sortM) = entryDates(sortN)
                            entryDates(sortN) = tmpDate
                            tmpHour = entryHours(sortM)
                            entryHours(sortM) = entryHours(sortN)
                            entryHours(sortN) = tmpHour
                        End If
                    Next sortN
                Next sortM

                ' Write sorted entries to the report
                For entryIdx = 1 To entryCount
                    studyHourValue = entryHours(entryIdx)
                    dateValue = entryDates(entryIdx)

                    Set templateRowRange = wsTemplate.Range("B12:F12")
                    templateRowRange.Copy
                    wsNew.Range("B" & lineNumber).PasteSpecial Paste:=xlPasteAll

                    If IsDate(dateValue) Then
                        wsNew.Range("B" & lineNumber).value = Format(CDate(dateValue), "dd.mm.yyyy")
                    Else
                        wsNew.Range("B" & lineNumber).value = "Invalid Date"
                    End If

                    wsNew.Range("D" & lineNumber).value = studyHourValue

                    calculatedValueC = Application.WorksheetFunction.Round(studyHourValue * 45, 0)
                    wsNew.Range("C" & lineNumber).value = calculatedValueC

                    wsNew.Range("E" & lineNumber).value = costPerHour

                    If IsNumeric(costPerHour) And IsNumeric(studyHourValue) Then
                        calculatedValueF = WorksheetFunction.Round(costPerHour * studyHourValue, 2)
                        wsNew.Range("F" & lineNumber).value = calculatedValueF
                    Else
                        wsNew.Range("F" & lineNumber).value = "N/A"
                    End If

                    wsNew.Range("B" & lineNumber).NumberFormat = "dd.mm.yyyy"

                    totalCostAllDisciplines = totalCostAllDisciplines + calculatedValueF
                    totalCostFromRecord = totalCostFromRecord + calculatedValueF
                    totalHoursFromRecord = totalHoursFromRecord + Round(studyHourValue, 2)

                    lineNumber = lineNumber + 1
                Next entryIdx
            End If

            ' Footer row with totals for this discipline group
            Set templateFooterRange = wsTemplate.Range("B14:F14")
            templateFooterRange.Copy
            wsNew.Range("B" & lineNumber).PasteSpecial Paste:=xlPasteAll

            wsNew.Range("F" & lineNumber).value = totalCostFromRecord
            wsNew.Range("F" & lineNumber).NumberFormat = "0.00"
            wsNew.Range("D" & lineNumber).value = totalHoursFromRecord
            wsNew.Range("D" & lineNumber).NumberFormat = "0.00"

            lineNumber = lineNumber + 1
            lineNumber = lineNumber + 1 ' Empty row for visual separation
        Next discGroupKey
        
        ' After all tables for the child, insert two empty rows
        lineNumber = lineNumber + 2
        
        ' Copy the footer template from Template sheet (B17:F17) to target workbook
        Set templateFooterRange = wsTemplate.Range("B17:F17")
        templateFooterRange.Copy
        wsNew.Range("B" & lineNumber).PasteSpecial Paste:=xlPasteAll
        
        ' Populate the total cost in FlineNumber with the sum of AQ cells, rounded to two decimals
        wsNew.Range("F" & lineNumber).value = WorksheetFunction.Round(totalCostAllDisciplines, 2)
        wsNew.Range("F" & lineNumber).NumberFormat = "0.00"
        
        ' Increment lineNumber after footer
        lineNumber = lineNumber + 1
        
        ' Replace any invalid characters in file name
        
        If IsDate(wsNew.Range("F8").value) Then
            reportDate = wsNew.Range("F8").value
        Else
            ' If F8 is not a valid date, default to current date
            reportDate = Date
        End If
        
        monthNumber = Month(reportDate)
        yearNumber = Year(reportDate)
        
        ' Define base file name
        baseFileName = childLastName & "_" & childFirstName & "_" & monthNumber & "_" & yearNumber
        
        ' Replace invalid characters in file name
        baseFileName = ReplaceInvalidFileNameChars(baseFileName)
        
        ' Define Excel and PDF file names
        fileNameXLSX = baseFileName & ".xlsx"
        fileNamePDF = baseFileName & ".pdf"
        
        ' Define the full paths for Excel and PDF
        excelPath = fso.BuildPath(subfolderPath, fileNameXLSX)
        pdfPath = fso.BuildPath(subfolderPath, fileNamePDF)
        
        ' Save the workbook as Excel file
        'On Error GoTo SaveExcelError
        'wbNew.SaveAs fileName:=excelPath, FileFormat:=xlOpenXMLWorkbook
        'On Error GoTo 0
        
        ' Export the report as PDF
        On Error GoTo ExportError
        wbNew.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        On Error GoTo 0
        
        ' Close the new workbook without saving (already saved as Excel)
        wbNew.Close SaveChanges:=False
        GoTo NextChild
        
SaveExcelError:
        MsgBox "An error occurred while saving the Excel report for " & childLastName & " " & childFirstName & "." & vbCrLf & _
            "Error: " & Err.Description, vbCritical, "Save Excel Error"
        ' Close the new workbook without saving
        If Not wbNew Is Nothing Then
            wbNew.Close SaveChanges:=False
        End If
        Resume NextChild
        
ExportError:
        MsgBox "An error occurred while exporting the PDF report for " & childLastName & " " & childFirstName & "." & vbCrLf & _
            "Error: " & Err.Description, vbCritical, "Export PDF Error"
        ' Close the new workbook without saving
        If Not wbNew Is Nothing Then
            wbNew.Close SaveChanges:=False
        End If
        Resume NextChild
        
NextChild:
    Next childKey
    
    ' Finalize progress bar
    UpdateProgressBar 100
    frmProgress.lblStatus.Caption = "Report generation completed."
    DoEvents ' Allow the form to update
    Application.Wait Now + TimeValue("0:00:02") ' Wait for 2 seconds to show completion
    Unload frmProgress
    
    ' Inform the user that reports have been generated
    MsgBox "Reports have been successfully generated and saved to:" & vbCrLf & reportFolder, vbInformation, "Operation Completed"
    
    ' Restore Excel settings
CleanUp:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
End Sub

' Helper function to replace invalid characters in file names
Function ReplaceInvalidFileNameChars(fileName As String) As String
    Dim invalidChars As Variant
    Dim ch As Variant
    
    invalidChars = Array("/", "\", ":", "*", "?", """", "<", ">", "|")
    
    For Each ch In invalidChars
        fileName = Replace(fileName, ch, "_")
    Next ch
    
    ReplaceInvalidFileNameChars = fileName
End Function

' Subroutine to update the progress bar based on percentage
Sub UpdateProgressBar(percent As Integer)
    With frmProgress
        ' Ensure percent is between 0 and 100
        If percent < 0 Then percent = 0
        If percent > 100 Then percent = 100
        
        ' Calculate the new width for lblProgress
        Dim frameWidth As Single
        frameWidth = .fraProgress.Width
        
        .lblProgress.Width = (percent / 100) * frameWidth
        
        ' Update percentage display
        .lblStatus.Caption = "Progress: " & percent & "%"
    End With
End Sub

' Subroutine to log skipped rows due to missing data
Sub LogSkippedRow(RowNumber As Long)
    Dim wsErrorLog As Worksheet
    On Error Resume Next
    Set wsErrorLog = ThisWorkbook.Sheets("ErrorLog")
    On Error GoTo 0
    If wsErrorLog Is Nothing Then
        Set wsErrorLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsErrorLog.name = "ErrorLog"
        wsErrorLog.Range("A1").value = "Skipped Rows Due to Missing Data"
        wsErrorLog.Range("A2").value = "Row Number"
    End If
    wsErrorLog.Range("A" & wsErrorLog.rows.Count).End(xlUp).Offset(1, 0).value = RowNumber
End Sub




