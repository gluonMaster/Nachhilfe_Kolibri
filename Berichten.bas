Attribute VB_Name = "Berichten"
Option Explicit

Sub GenerateChildReportsWithDetailedTables()
    ' Define variables
    Dim wsLoad As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsTemplate2 As Worksheet  ' New variable for second template
    Dim currentTemplate As Worksheet ' Variable to track which template to use for current record
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
    Dim useTemplate2 As Boolean  ' Flag to indicate which template to use
    
    ' Initialize
    Set wsLoad = ThisWorkbook.ActiveSheet ' Assumes the user is on the monthly load sheet
    Set wsTemplate = ThisWorkbook.Sheets("Shablon") ' Template sheet
    Set wsTemplate2 = ThisWorkbook.Sheets("Shablon2") ' Second template sheet
    Set dictChildren = CreateObject("Scripting.Dictionary") ' Late binding
    
    ' Determine the last row with data in column A (Child ID)
    lastRow = wsLoad.Cells(wsLoad.Rows.Count, "A").End(xlUp).Row
    If lastRow < 11 Then
        MsgBox "No data found starting from row 11.", vbExclamation
        Exit Sub
    End If
    
    ' Ask the user if they want to process only selected rows
    If MsgBox("Do you want to generate reports only for the selected records?", vbYesNo + vbQuestion, "Generate Reports") = vbYes Then
        processSelected = True
        ' Check if there are selected rows
        If TypeName(Selection) <> "Range" Then
            MsgBox "Please select the rows for which you want to generate reports.", vbExclamation
            processSelected = False
        Else
            ' Check that only entire rows are selected and they start from row 11
            Set rowsToProcess = Nothing
            For Each selectedRow In Selection.Rows
                If selectedRow.Row < 11 Or selectedRow.Row > lastRow Then
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
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Destination Folder for Reports"
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
        For Each selectedRow In rowsToProcess.Rows
            currentRow = selectedRow.Row
            ' Read necessary cells
            Dim cellA As String
            Dim cellB As String
            Dim cellC As String
            Dim cellD As String
            Dim cellE As String
            Dim cellF As Variant
            Dim cellG As Variant
            Dim cellH As Variant
            Dim cellI As Variant  ' Added for template selection
            Dim cellAU As String
            Dim cellAV As Variant
            Dim cellAP As Double
            Dim cellAQ As Double
            
            cellA = Trim(wsLoad.Cells(currentRow, "A").Value) ' Child ID
            cellB = Trim(wsLoad.Cells(currentRow, "B").Value) ' Last Name
            cellC = Trim(wsLoad.Cells(currentRow, "C").Value) ' First Name
            cellD = Trim(wsLoad.Cells(currentRow, "D").Value) ' Discipline
            cellE = Trim(wsLoad.Cells(currentRow, "E").Value) ' Lesson Type
            cellF = wsLoad.Cells(currentRow, "F").Value ' Start Date
            cellG = wsLoad.Cells(currentRow, "G").Value ' End Date
            cellH = wsLoad.Cells(currentRow, "H").Value ' Age
            cellI = wsLoad.Cells(currentRow, "I").Value ' Check for template selection
            cellAU = Trim(wsLoad.Cells(currentRow, "AU").Value) ' Social Service ID
            cellAV = wsLoad.Cells(currentRow, "AV").Value ' Birth Date
            cellAP = wsLoad.Cells(currentRow, "AP").Value ' Cost per Hour
            cellAQ = wsLoad.Cells(currentRow, "AQ").Value ' Total Cost
            
            ' Skip rows with missing critical data
            If cellA = "" Or cellB = "" Or cellC = "" Or cellD = "" Or cellE = "" Or cellF = "" Or cellG = "" Or cellH = "" Then
                ' Log skipped rows
                Call LogSkippedRow(currentRow)
                GoTo NextSelectedRow
            End If
            
            ' Determine which template to use based on cellI
            useTemplate2 = (cellI = 2)
            
            ' Create a unique key for each child
            childKey = cellA & "|" & cellB & "|" & cellC & "|" & cellF & "|" & cellG & "|" & cellH
            
            ' If the child is not yet in the dictionary, add them with a new collection
            If Not dictChildren.Exists(childKey) Then
                Set recordCollection = New Collection
                dictChildren.Add childKey, recordCollection
            Else
                Set recordCollection = dictChildren(childKey)
            End If
            
            ' Add the current record to the child's collection
            ' Record data includes: Discipline, Lesson Type, Cost per Hour, Total Cost, Social Service ID, Birth Date, Row Number, Template Flag
            recordCollection.Add Array(cellD, cellE, cellAP, cellAQ, cellAU, cellAV, currentRow, useTemplate2)
        
NextSelectedRow:
        Next selectedRow
    Else
        ' Process all rows starting from row 11
        For currentRow = 11 To lastRow
            ' Read necessary cells
            
            cellA = Trim(wsLoad.Cells(currentRow, "A").Value) ' Child ID
            cellB = Trim(wsLoad.Cells(currentRow, "B").Value) ' Last Name
            cellC = Trim(wsLoad.Cells(currentRow, "C").Value) ' First Name
            cellD = Trim(wsLoad.Cells(currentRow, "D").Value) ' Discipline
            cellE = Trim(wsLoad.Cells(currentRow, "E").Value) ' Lesson Type
            cellF = wsLoad.Cells(currentRow, "F").Value ' Start Date
            cellG = wsLoad.Cells(currentRow, "G").Value ' End Date
            cellH = wsLoad.Cells(currentRow, "H").Value ' Age
            cellI = wsLoad.Cells(currentRow, "I").Value ' Check for template selection
            cellAU = Trim(wsLoad.Cells(currentRow, "AU").Value) ' Social Service ID
            cellAV = wsLoad.Cells(currentRow, "AV").Value ' Birth Date
            cellAP = wsLoad.Cells(currentRow, "AP").Value ' Cost per Hour
            cellAQ = wsLoad.Cells(currentRow, "AQ").Value ' Total Cost
            
            ' Skip rows with missing critical data
            If cellA = "" Or cellB = "" Or cellC = "" Or cellD = "" Or cellE = "" Or cellF = "" Or cellG = "" Or cellH = "" Then
                ' Log skipped rows
                Call LogSkippedRow(currentRow)
                GoTo NextRow
            End If
            
            ' Determine which template to use based on cellI
            useTemplate2 = (cellI = 2)
            
            ' Create a unique key for each child
            childKey = cellA & "|" & cellB & "|" & cellC & "|" & cellF & "|" & cellG & "|" & cellH
            
            ' If the child is not yet in the dictionary, add them with a new collection
            If Not dictChildren.Exists(childKey) Then
                Set recordCollection = New Collection
                dictChildren.Add childKey, recordCollection
            Else
                Set recordCollection = dictChildren(childKey)
            End If
            
            ' Add the current record to the child's collection
            ' Record data includes: Discipline, Lesson Type, Cost per Hour, Total Cost, Social Service ID, Birth Date, Row Number, Template Flag
            recordCollection.Add Array(cellD, cellE, cellAP, cellAQ, cellAU, cellAV, currentRow, useTemplate2)
        
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
        lessonStartDate = splitKey(3) ' Start Date
        lessonEndDate = splitKey(4) ' End Date
        ' Age is splitKey(5), but not used here
        
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
        
        ' Determine which template to use for this child based on the first record's template flag
        ' We use the template flag from the first record for the child
        useTemplate2 = dictChildren(childKey)(1)(7) ' 8th element (index 7) is the template flag
        
        ' Select the appropriate template
        If useTemplate2 Then
            Set currentTemplate = wsTemplate2 ' Use Shablon2
        Else
            Set currentTemplate = wsTemplate ' Use Shablon
        End If
        
        ' Copy the initial template range (A1:F9) from chosen Template to the new workbook        
        Set templateRange = currentTemplate.Range("A1:F9")
        templateRange.Copy
        wsNew.Range("A1").PasteSpecial Paste:=xlPasteAll
        
        ' Copy the column widths from chosen Template to the new worksheet
        currentTemplate.Columns("A:F").Copy
        wsNew.Columns("A:F").PasteSpecial Paste:=xlPasteColumnWidths
        
        ' Populate specific cells with child data
        Dim combinedName As String
        combinedName = childLastName & ", " & childFirstName
        wsNew.Range("C2").Value = combinedName
        wsNew.Range("E3").Value = dictChildren(childKey)(1)(5) ' Birth Date from first record
        wsNew.Range("C3").Value = dictChildren(childKey)(1)(4) ' Social Service ID from first record
        wsNew.Range("C7").Value = lessonStartDate ' Start Date
        wsNew.Range("C8").Value = lessonEndDate ' End Date

        If useTemplate2 Then
            wsNew.Range("B7").Value = wsLoad.Range("AW" & dictChildren(childKey)(1)(6)).Value ' Previous start Date
            wsNew.Range("B8").Value = wsLoad.Range("AX" & dictChildren(childKey)(1)(6)).Value ' Previous end Date
            wsNew.Range("B7").NumberFormat = "dd.mm.yyyy"
            wsNew.Range("B8").NumberFormat = "dd.mm.yyyy"
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
            foundRow = foundCell.Row
            ' Place the address parts into E4 and E5
            wsNew.Range("E4").Value = wsKinder.Cells(foundRow, 19).Value
            wsNew.Range("E5").Value = wsKinder.Cells(foundRow, 20).Value
            
            ' Format the cells E4 and E5
            With wsNew.Range("E4:E5")
                .HorizontalAlignment = xlLeft
                .Font.Bold = True
                .Font.Name = "Calibri"
                .Font.Size = 10
            End With
        Else
            ' If no match is found, you can leave these cells blank or handle it differently if needed
            wsNew.Range("E4").Value = ""
            wsNew.Range("E5").Value = ""
        End If
        
        ' -------------------------------------------------------
        
        ' Format the date cells as desired (e.g., dd.mm.yyyy)
        wsNew.Range("E3").NumberFormat = "dd.mm.yyyy"
        wsNew.Range("C7").NumberFormat = "dd.mm.yyyy"
        wsNew.Range("C8").NumberFormat = "dd.mm.yyyy"
        
        ' Initialize lineNumber for table entries
        ' Assuming that after A1:F12, the tables start from row 10
        lineNumber = 10
        
        ' Iterate through each record (discipline) of the child
        For i = 1 To dictChildren(childKey).Count
            ' Retrieve record data
            recordData = dictChildren(childKey)(i)
            disciplineName = recordData(0) ' Discipline
            lessonTypeCode = recordData(1) ' Lesson Type (G or I)
            costPerHour = recordData(2) ' Cost per Hour
            totalCostFromRecord = 0 ' Total Cost from AQ
            socialServiceID = recordData(4) ' Social Service ID
            childBirthDate = recordData(5) ' Birth Date
            currentRow = recordData(6) ' Row Number in source sheet
            useTemplate2 = recordData(7) ' Template flag
            
            ' Update the current template for this specific record if needed
            If useTemplate2 Then
                Set currentTemplate = wsTemplate2 ' Use Shablon2
            Else
                Set currentTemplate = wsTemplate ' Use Shablon
            End If
            
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
            
            ' Copy the table header from current Template sheet (B10:F11) to target workbook
            Set templateHeaderRange = currentTemplate.Range("B10:F11")
            templateHeaderRange.Copy
            wsNew.Range("B" & lineNumber).PasteSpecial Paste:=xlPasteAll
            
            ' Populate the header string into the appropriate cell (ClineNumber)
            wsNew.Range("C" & lineNumber).Value = headerString
            
            ' Move to the next line for table rows
            lineNumber = lineNumber + 2 ' Assuming header occupies 2 rows (B10:F11)
            
            ' Iterate through columns J to AU (10 to 40) for study hours
            Dim col As Long
            For col = 10 To 40 ' Columns J to AU
                studyHourValue = Round(wsLoad.Cells(currentRow, col).Value / 45, 2) ' Study hours for the day
                If IsNumeric(studyHourValue) Then
                    If studyHourValue > 0 Then
                        ' Copy the table row template from current Template sheet (B12:F12) to target workbook
                        Set templateRowRange = currentTemplate.Range("B12:F12")
                        templateRowRange.Copy
                        wsNew.Range("B" & lineNumber).PasteSpecial Paste:=xlPasteAll
                        
                        ' Populate the table row
                        ' BlineNumber: Date from row 5 of the current column
                        dateValue = wsLoad.Cells(5, col).Value
                        If IsDate(dateValue) Then
                            wsNew.Range("B" & lineNumber).Value = Format(CDate(dateValue), "dd.mm.yyyy")
                        Else
                            wsNew.Range("B" & lineNumber).Value = "Invalid Date"
                        End If
                        
                        ' DlineNumber: Hours from current cell
                        wsNew.Range("D" & lineNumber).Value = studyHourValue
                        
                        ' ClineNumber: 45 * Hours, rounded to integer
                        If IsNumeric(studyHourValue) Then
                            calculatedValueC = Application.WorksheetFunction.Round(studyHourValue * 45, 0)
                            wsNew.Range("C" & lineNumber).Value = calculatedValueC
                        Else
                            wsNew.Range("C" & lineNumber).Value = "N/A"
                        End If
                        
                        ' ElineNumber: Cost per hour from AP
                        wsNew.Range("E" & lineNumber).Value = costPerHour
                        
                        ' FlineNumber: E * D, rounded to two decimal places
                        If IsNumeric(costPerHour) And IsNumeric(studyHourValue) Then
                            calculatedValueF = WorksheetFunction.Round(costPerHour * studyHourValue, 2)
                            wsNew.Range("F" & lineNumber).Value = calculatedValueF
                        Else
                            wsNew.Range("F" & lineNumber).Value = "N/A"
                        End If
                        
                        ' Format the date cell
                        wsNew.Range("B" & lineNumber).NumberFormat = "dd.mm.yyyy"
                        
                        ' Accumulate total cost for all disciplines
                        totalCostAllDisciplines = totalCostAllDisciplines + calculatedValueF
                        totalCostFromRecord = totalCostFromRecord + calculatedValueF
                        
                        ' Increment lineNumber for next entry
                        lineNumber = lineNumber + 1
                    End If
                End If
            Next col
            
            ' After processing all study hours, copy the footer row from current Template sheet (B14:F14)
            Set templateFooterRange = currentTemplate.Range("B14:F14")
            templateFooterRange.Copy
            wsNew.Range("B" & lineNumber).PasteSpecial Paste:=xlPasteAll
            
            ' Populate the total cost in FlineNumber
            wsNew.Range("F" & lineNumber).Value = totalCostFromRecord
            wsNew.Range("F" & lineNumber).NumberFormat = "0.00"
            
            ' Increment lineNumber after footer
            lineNumber = lineNumber + 1
            
            ' Insert an empty row for better visual separation between tables
            lineNumber = lineNumber + 1
        Next i
        
        ' After all tables for the child, insert two empty rows
        lineNumber = lineNumber + 2
        
        ' Use the template that was selected for the last discipline
        If useTemplate2 Then
            Set currentTemplate = wsTemplate2 ' Use Shablon2
        Else
            Set currentTemplate = wsTemplate ' Use Shablon
        End If
        
        ' Copy the footer template from current Template sheet (B17:F17) to target workbook
        Set templateFooterRange = currentTemplate.Range("B17:F17")
        templateFooterRange.Copy
        wsNew.Range("B" & lineNumber).PasteSpecial Paste:=xlPasteAll
        
        ' Populate the total cost in FlineNumber with the sum of AQ cells, rounded to two decimals
        wsNew.Range("F" & lineNumber).Value = WorksheetFunction.Round(totalCostAllDisciplines, 2)
        wsNew.Range("F" & lineNumber).NumberFormat = "0.00"
        
        ' Increment lineNumber after footer
        lineNumber = lineNumber + 1
        
        ' Replace any invalid characters in file name
        
        If IsDate(wsNew.Range("F8").Value) Then
            reportDate = wsNew.Range("F8").Value
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
        'wbNew.SaveAs Filename:=excelPath, FileFormat:=xlOpenXMLWorkbook
        'On Error GoTo 0
        
        ' Export the report as PDF
        On Error GoTo ExportError
        wbNew.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
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
Cleanup:
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
Sub LogSkippedRow(rowNumber As Long)
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
    wsErrorLog.Range("A" & wsErrorLog.Rows.Count).End(xlUp).Offset(1, 0).value = rowNumber
End Sub




