Attribute VB_Name = "MonatTabelleEinfullung"
Option Explicit

Sub PopulateCurrentSheetWithExceptions()
    On Error GoTo ErrorHandler
    
    ' Declare necessary variables
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim referenceDate As Date
    Dim sourceLastRow As Long
    Dim targetLastRow As Long
    Dim i As Long
    Dim copyFromGH As Boolean
    Dim copyFromM As Boolean
    Dim newRecordKey As String
    Dim dictDuplicates As Object
    Dim dictMain As Object
    Dim existingRow As Range
    Dim sortRange As Range
    
    ' Initialize the dictionaries
    Set dictDuplicates = CreateObject("Scripting.Dictionary")
    Set dictMain = CreateObject("Scripting.Dictionary")
    
    ' **1. Set the source and target worksheets**
    Set sourceWs = ThisWorkbook.Sheets("Kinder")
    Set targetWs = ActiveSheet ' Assumes the macro is run from the target sheet
    
    If ClearAllFilters(sourceWs) And ClearAllFilters(targetWs) Then
        ' **2. Read the reference date from A1 of the target sheet**
        If IsDate(targetWs.Range("A1").value) Then
            referenceDate = targetWs.Range("A1").value
        Else
            MsgBox "The value in A1 is not a valid date. Please enter a valid date.", vbExclamation
            Exit Sub
        End If
        
        ' **3. Determine the last row in the source sheet**
        sourceLastRow = sourceWs.Cells(sourceWs.Rows.Count, "B").End(xlUp).row
        
        ' **4. Determine the last filled row in the target sheet**
        targetLastRow = targetWs.Cells(targetWs.Rows.Count, "A").End(xlUp).row
        If targetLastRow < 11 Then
            targetLastRow = 10 ' Data will start from row 11
        End If
        
        ' **5. Populate the dictionaries with existing records**
        If targetLastRow >= 11 Then
            For Each existingRow In targetWs.Range("A11:A" & targetLastRow).Rows
                ' **a. Duplicate Key: A, B, C, D, E, F, G, H, AU, AV**
                Dim dupKey As String
                dupKey = existingRow.Cells(1, "A").value & "|" & _
                         existingRow.Cells(1, "B").value & "|" & _
                         existingRow.Cells(1, "C").value & "|" & _
                         existingRow.Cells(1, "D").value & "|" & _
                         existingRow.Cells(1, "E").value & "|" & _
                         existingRow.Cells(1, "F").value & "|" & _
                         existingRow.Cells(1, "G").value & "|" & _
                         existingRow.Cells(1, "H").value & "|" & _
                         existingRow.Cells(1, "AU").value & "|" & _
                         existingRow.Cells(1, "AV").value
                         
                ' Add to duplicates dictionary
                If Not dictDuplicates.exists(dupKey) Then
                    dictDuplicates.Add dupKey, True
                End If
                
                ' **b. Main Key: A, B, C, D, F, G, H**
                ' Corresponds to source columns B, C, D, E, G, H, M
                Dim mainKey As String
                mainKey = existingRow.Cells(1, "A").value & "|" & _
                          existingRow.Cells(1, "B").value & "|" & _
                          existingRow.Cells(1, "C").value & "|" & _
                          existingRow.Cells(1, "D").value & "|" & _
                          existingRow.Cells(1, "F").value & "|" & _
                          existingRow.Cells(1, "G").value & "|" & _
                          existingRow.Cells(1, "H").value        ' H: source M
                          
                ' Store the E field value (source F)
                If Not dictMain.exists(mainKey) Then
                    dictMain.Add mainKey, existingRow.Cells(1, "E").value
                Else
                    ' If multiple entries exist, ensure that if any E is not empty, it's recorded
                    If dictMain(mainKey) = "" And existingRow.Cells(1, "E").value <> "" Then
                        dictMain(mainKey) = existingRow.Cells(1, "E").value
                    End If
                End If
            Next existingRow
        End If
        
        ' **6. Improve performance**
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
        ' **7. Loop through each row in the source sheet starting from row 5**
        For i = 5 To sourceLastRow
            ' **a. Read relevant data from source**
            Dim sourceDataB As Variant, sourceDataC As Variant, sourceDataD As Variant
            Dim sourceDataE As Variant, sourceDataF As Variant, sourceDataG As Variant
            Dim sourceDataH As Variant, sourceDataM As Variant, sourceDataO As Variant
            Dim sourceDataL As Variant
            Dim sourceDateH As Variant
            Dim sourceDataK As Variant
            
            sourceDataB = sourceWs.Cells(i, "B").value
            sourceDataC = sourceWs.Cells(i, "C").value
            sourceDataD = sourceWs.Cells(i, "D").value
            sourceDataE = sourceWs.Cells(i, "E").value
            sourceDataF = sourceWs.Cells(i, "F").value
            sourceDataG = sourceWs.Cells(i, "G").value
            sourceDataH = sourceWs.Cells(i, "H").value
            sourceDataM = sourceWs.Cells(i, "M").value
            sourceDataO = sourceWs.Cells(i, "O").value
            sourceDataL = sourceWs.Cells(i, "L").value
            sourceDataK = sourceWs.Cells(i, "K").value
            sourceDataK = Replace(LCase(Trim(sourceDataK)), " ", "")
            
            sourceDateH = sourceDataH
            
            ' **b. Initialize flags for copying**
            copyFromGH = False
            copyFromM = False
            
            ' **c. Check if H date is greater than reference date**
            If IsDate(sourceDateH) Then
                If CDate(sourceDateH) > referenceDate Then
                    ' **i. Check if G date exceeds reference date**
                    Dim sourceDateG As Variant
                    sourceDateG = sourceDataG
                    If IsDate(sourceDateG) Then
                        If CDate(sourceDateG) <= referenceDate Then
                            copyFromGH = True
                        Else
                            ' If G > referenceDate, check if G is in the same month and year
                            If Month(CDate(sourceDateG)) = Month(referenceDate) And _
                               Year(CDate(sourceDateG)) = Year(referenceDate) Then
                                copyFromGH = True
                            Else
                                copyFromGH = False
                            End If
                        End If
                    Else
                        ' If G is not a valid date, proceed to copy
                        copyFromGH = True
                    End If
                End If
            End If
            
            ' **d. Perform copying based on flags**
            ' **Copy from B to H if applicable**
            If copyFromGH Then
                ' **Create a unique key for the new record (for duplicate checking)**
                newRecordKey = sourceDataB & "|" & _
                              sourceDataC & "|" & _
                              sourceDataD & "|" & _
                              sourceDataE & "|" & _
                              sourceDataF & "|" & _
                              sourceDataG & "|" & _
                              sourceDataH & "|" & _
                              sourceDataM & "|" & _
                              sourceDataO & "|" & _
                              sourceDataL
                              
                ' **Create main key based on A, B, C, D, F, G, H**
                ' Corresponds to source B, C, D, E, G, H, M
                Dim newMainKey As String
                newMainKey = sourceDataB & "|" & _
                             sourceDataC & "|" & _
                             sourceDataD & "|" & _
                             sourceDataE & "|" & _
                             sourceDataG & "|" & _
                             sourceDataH & "|" & _
                             sourceDataM
                             
                ' **Check duplicate**
                If Not dictDuplicates.exists(newRecordKey) Then
                    ' **Check the additional condition**
                    If dictMain.exists(newMainKey) Then
                        If IsEmpty(sourceDataF) Or Trim(sourceDataF) = "" Then
                            If Trim(dictMain(newMainKey)) <> "" Then
                                ' Skip adding this record
                                GoTo NextIteration
                            End If
                        End If
                    End If
                    
                    ' **Add the record to the target sheet**
                    targetLastRow = targetLastRow + 1
                    targetWs.Cells(targetLastRow, "A").value = sourceDataB ' B > A
                    targetWs.Cells(targetLastRow, "B").value = sourceDataC ' C > B
                    targetWs.Cells(targetLastRow, "C").value = sourceDataD ' D > C
                    targetWs.Cells(targetLastRow, "D").value = sourceDataE ' E > D
                    targetWs.Cells(targetLastRow, "E").value = sourceDataF ' F > E
                    targetWs.Cells(targetLastRow, "F").value = sourceDataG ' G > F
                    targetWs.Cells(targetLastRow, "G").value = sourceDataH ' H > G
                    targetWs.Cells(targetLastRow, "H").value = sourceDataM ' M > H
                    targetWs.Cells(targetLastRow, "AU").value = sourceDataO ' O > AU
                    targetWs.Cells(targetLastRow, "AV").value = sourceDataL ' L > AV
                    
                    If sourceDataK <> "jobcenter" Then
                         Dim fillColor As Long
                         fillColor = RGB(255, 102, 102)
                         Range("I" & targetLastRow).Interior.Color = fillColor
                         targetWs.Cells(targetLastRow, "I").value = "!"
                    End If
                    
                    ' **Add the new record key to the duplicates dictionary to prevent future duplicates**
                    dictDuplicates.Add newRecordKey, True
                    
                    ' **Update the main dictionary**
                    If Not dictMain.exists(newMainKey) Then
                        dictMain.Add newMainKey, sourceDataF
                    Else
                        ' If existing E is empty and new E is not, update it
                        If Trim(dictMain(newMainKey)) = "" And Not (IsEmpty(sourceDataF) Or Trim(sourceDataF) = "") Then
                            dictMain(newMainKey) = sourceDataF
                        End If
                    End If
                End If
            End If
            
NextIteration:
        Next i
        
        ' **8. Sort the data from row 11 onwards based on column B**
        If targetLastRow >= 11 Then
            Set sortRange = targetWs.Range("A11:AV" & targetLastRow)
            With targetWs.Sort
                .SortFields.Clear
                .SortFields.Add key:=targetWs.Range("B11:B" & targetLastRow), Order:=xlAscending ' sortkey (for B)
                .SortFields.Add key:=targetWs.Range("C11:C" & targetLastRow), Order:=xlAscending ' sortkey (for C)
                .SetRange sortRange
                .Header = xlNo
                .Apply
            End With
            
            'With sortRange
                '.Sort Key1:=targetWs.Range("B11"), Order1:=xlAscending, Header:=xlNo
                '.Sort Key2:=targetWs.Range("C11"), Order1:=xlAscending, Header:=xlNo
            'End With
        End If
        
        Call ApplyDateBasedFormatting(targetWs)
        
        ' **9. Restore settings**
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        
        MsgBox "Data population with exceptions completed successfully.", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


Public Function ClearAllFilters(ws As Worksheet) As Boolean
    ' This function removes all active filters on the specified worksheet
    ' Returns True if filters were cleared successfully, False if there was an error
    '
    ' Parameters:
    ' ws - The worksheet to clear filters from
    
    On Error GoTo ErrorHandler
    
    ' Check if the worksheet has any filters
    If ws.FilterMode Then
        ' Clear all filters
        ws.ShowAllData
        ClearAllFilters = True
    Else
        ' No filters to clear
        ClearAllFilters = True
    End If
    
    Exit Function
    
ErrorHandler:
    ' Return False if there was an error
    ClearAllFilters = False
End Function

Sub ApplyDateBasedFormatting(ws As Worksheet)
    ' This procedure clears existing conditional formatting and applies new formatting
    ' based on date ranges and weekday criteria
    '
    ' Parameters:
    ' ws - The worksheet to apply formatting to
    
    Dim lastRow As Long
    Dim formatRange As Range
    Dim i As Long, j As Long
    Dim cellRange As Range
    Dim columnDate As Date
    Dim allowedStartDate As Date
    Dim allowedEndDate As Date
    
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Exit if there are no rows to process
    If lastRow < 11 Then
        MsgBox "No data found to format", vbInformation
        GoTo CleanExit
    End If
    
    ' Define the format range
    Set formatRange = ws.Range("J11:AN" & lastRow)
    
    ' Clear any existing conditional formatting
    formatRange.FormatConditions.Delete
    
    ' Process each cell individually
    For i = 11 To lastRow
        ' Get allowed date range for this row
        On Error Resume Next
        allowedStartDate = ws.Cells(i, "F").value
        allowedEndDate = ws.Cells(i, "G").value
        On Error GoTo 0
        
        ' Skip row if date range is invalid
        If Not (IsDate(allowedStartDate) And IsDate(allowedEndDate)) Then
            GoTo NextRow
        End If
        
        ' Process each column in the row
        For j = 10 To 40  ' Columns J to AN (10 to 40)
            Set cellRange = ws.Cells(i, j)
            
            ' Get the date from row 5 for this column
            On Error Resume Next
            columnDate = ws.Cells(5, j).value
            On Error GoTo 0
            
            If IsDate(columnDate) Then
                ' Apply formatting based on criteria
                If Weekday(columnDate) = 1 Then  ' Sunday
                    cellRange.Interior.Color = RGB(247, 180, 65)  ' Coffee color
                ElseIf Weekday(columnDate) = 7 Then  ' Saturday
                    cellRange.Interior.Color = RGB(253, 253, 142)  ' Yellow
                ElseIf columnDate >= allowedStartDate And columnDate <= allowedEndDate Then
                    cellRange.Interior.Color = RGB(144, 238, 144)  ' Light green
                Else
                    cellRange.Interior.ColorIndex = xlNone  ' No color
                End If
            End If
        Next j
        
NextRow:
    Next i
    
    MsgBox "Formatting applied successfully", vbInformation
    
CleanExit:
    ' Restore settings
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub
