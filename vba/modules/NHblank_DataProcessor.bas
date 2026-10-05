Attribute VB_Name = "NHblank_DataProcessor"
Option Explicit

Private Const PROVIDER_NAME As String = _
    "Kinder- und Elternzentrum Kolibri e.V."

' False: Sozialamt/non-Jobcenter records create only the legacy XLSX blank.
' True:  restore the previous behavior and create both DOCX and XLSX.
Private Const GENERATE_WORD_FOR_NON_JOBCENTER As Boolean = False

' Process records and create monthly blanks.
' Standard Jobcenter number: one DOCX.
' Sozialamt number: one legacy XLSX; optional DOCX is controlled above.
' Returns the number of files created.
Public Function ProcessRecords( _
    ByVal processAllRecords As Boolean, _
    ByVal wordTemplatePath As String, _
    ByVal legacyTemplatePath As String, _
    ByVal targetFolder As String, _
    ByRef activeRecordsFound As Long, _
    ByVal selectedRows As Collection, _
    ByVal overwriteExisting As Boolean, _
    ByRef processWarnings As String _
) As Long

    Dim wsKinder As Worksheet
    Dim wordApp As Object
    Dim lastRow As Long
    Dim currentRow As Long
    Dim selectedRow As Variant
    Dim filesCreated As Long
    Dim createdForRecord As Long
    Dim processError As String

    filesCreated = 0
    activeRecordsFound = 0
    processWarnings = ""

    On Error GoTo FatalError

    ' The generation source remains Kinder_Blanks.
    Set wsKinder = NHblank_Sheets.NHblank_EnsureKinderBlanksSheetExists()

    If wsKinder Is Nothing Then
        NHblank_UI.ShowError _
            "FEHLER: Worksheet 'Kinder_Blanks' konnte nicht erstellt werden!" & _
            vbCrLf & "Stellen Sie sicher, dass das Worksheet 'Kinder' existiert."
        Exit Function
    End If

    If processAllRecords Then
        lastRow = wsKinder.Cells(wsKinder.Rows.Count, "C").End(xlUp).Row

        For currentRow = 5 To lastRow
            If IsRecordActive(wsKinder, currentRow) Then
                activeRecordsFound = activeRecordsFound + 1
                createdForRecord = 0
                processError = ""

                If Not ProcessSingleRecord( _
                    wsKinder, wordApp, currentRow, _
                    wordTemplatePath, legacyTemplatePath, targetFolder, _
                    overwriteExisting, createdForRecord, processError) Then
                    AppendWarning processWarnings, currentRow, processError
                End If

                filesCreated = filesCreated + createdForRecord
            End If
        Next currentRow
    Else
        If selectedRows Is Nothing Then
            processWarnings = _
                "Interner Fehler: Die Zeilenauswahl wurde nicht uebergeben."
            GoTo CleanExit
        End If

        If selectedRows.Count = 0 Then
            processWarnings = "Keine ausgewaehlten Zeilen zum Verarbeiten."
            GoTo CleanExit
        End If

        For Each selectedRow In selectedRows
            currentRow = CLng(selectedRow)
            If currentRow > 4 Then
                If IsRecordActive(wsKinder, currentRow) Then
                    activeRecordsFound = activeRecordsFound + 1
                    createdForRecord = 0
                    processError = ""

                    If Not ProcessSingleRecord( _
                        wsKinder, wordApp, currentRow, _
                        wordTemplatePath, legacyTemplatePath, targetFolder, _
                        overwriteExisting, createdForRecord, _
                        processError) Then
                        AppendWarning _
                            processWarnings, currentRow, processError
                    End If

                    filesCreated = filesCreated + createdForRecord
                End If
            End If
        Next selectedRow
    End If

CleanExit:
    On Error Resume Next
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordApp = Nothing
    On Error GoTo 0

    ProcessRecords = filesCreated
    Exit Function

FatalError:
    AppendWarning processWarnings, currentRow, _
                  "Unerwarteter Fehler #" & CStr(Err.Number) & ": " & _
                  Err.Description
    Resume CleanExit
End Function

' Debug function to check font color of current cell.
Public Sub CheckCurrentCellColor()
    Dim msg As String
    Dim cell As Range
    Dim themeColorStr As String

    Set cell = ActiveCell

    On Error Resume Next
    themeColorStr = CStr(cell.Font.ThemeColor)
    If Err.Number <> 0 Then
        themeColorStr = "Error reading ThemeColor"
        Err.Clear
    End If
    On Error GoTo 0

    msg = "Color Debug for cell " & cell.Address & ":" & vbCrLf & vbCrLf
    msg = msg & "Font.ColorIndex: " & cell.Font.ColorIndex & vbCrLf
    msg = msg & "Font.Color: " & cell.Font.Color & vbCrLf
    msg = msg & "Font.ThemeColor: " & themeColorStr & vbCrLf
    msg = msg & "Cell Value: " & cell.Value & vbCrLf

    MsgBox msg, vbInformation, "Cell Color Info"
End Sub

Private Function IsRecordActive(ByVal ws As Worksheet, _
                                ByVal rowNum As Long) As Boolean
    Dim cell As Range

    Set cell = ws.Cells(rowNum, "C")

    If Len(Trim$(CStr(cell.Value))) = 0 Then Exit Function
    IsRecordActive = (cell.Font.ColorIndex <> 15)
End Function

Private Function ProcessSingleRecord( _
    ByVal wsKinder As Worksheet, _
    ByRef wordApp As Object, _
    ByVal rowNum As Long, _
    ByVal wordTemplatePath As String, _
    ByVal legacyTemplatePath As String, _
    ByVal targetFolder As String, _
    ByVal overwriteExisting As Boolean, _
    ByRef createdCount As Long, _
    ByRef errorMsg As String _
) As Boolean

    Dim lastName As String
    Dim firstName As String
    Dim studentName As String
    Dim discipline As String
    Dim teacherName As String
    Dim teacherFolder As String
    Dim targetSubfolder As String
    Dim birthDateValue As Variant
    Dim birthDateText As String
    Dim dateFrom As Variant
    Dim dateTo As Variant
    Dim referenceDate As Variant
    Dim bgNumber As String
    Dim isJobcenter As Boolean
    Dim bgError As String
    Dim fileName As String
    Dim wordOutputPath As String
    Dim legacyOutputPath As String
    Dim operationError As String
    Dim expectedCount As Long
    Dim createWordBlank As Boolean
    Dim fso As Object

    createdCount = 0
    errorMsg = ""
    expectedCount = 0

    On Error GoTo ErrorHandler
    Set fso = CreateObject("Scripting.FileSystemObject")

    lastName = Trim$(CStr(wsKinder.Cells(rowNum, "C").Value))
    firstName = Trim$(CStr(wsKinder.Cells(rowNum, "D").Value))
    discipline = Trim$(CStr(wsKinder.Cells(rowNum, "E").Value))
    teacherName = Trim$(CStr(wsKinder.Cells(rowNum, "I").Value))
    studentName = Trim$(lastName & " " & firstName)

    birthDateValue = wsKinder.Cells(rowNum, "L").Value
    dateFrom = wsKinder.Cells(rowNum, "G").Value
    dateTo = wsKinder.Cells(rowNum, "H").Value
    referenceDate = wsKinder.Range("T2").Value

    If Len(lastName) = 0 Then
        errorMsg = "Nachname (Spalte C) ist leer."
        Exit Function
    End If

    If Len(firstName) = 0 Then
        errorMsg = "Vorname (Spalte D) ist leer."
        Exit Function
    End If

    If Len(discipline) = 0 Then
        errorMsg = "Fach (Spalte E) ist leer."
        Exit Function
    End If

    If Not IsDate(birthDateValue) Then
        errorMsg = "Geburtsdatum (Spalte L) ist leer oder ungueltig."
        Exit Function
    End If
    birthDateText = Format$(CDate(birthDateValue), "dd.mm.yyyy")

    If Not IsDate(referenceDate) Then
        errorMsg = "Referenzdatum in Kinder_Blanks!T2 ist ungueltig."
        Exit Function
    End If

    If Not NHblank_BgNummer.NHblank_NormalizeAndProtectBgCell( _
        wsKinder.Cells(rowNum, "O"), _
        bgNumber, isJobcenter, bgError) Then
        errorMsg = bgError
        Exit Function
    End If

    createWordBlank = isJobcenter Or GENERATE_WORD_FOR_NON_JOBCENTER

    If Len(teacherName) = 0 Then
        teacherFolder = "Unsorted"
    Else
        teacherFolder = NHblank_Utils.CreateSafeFolderName(teacherName)
    End If

    targetSubfolder = targetFolder & "\" & teacherFolder
    If Not fso.FolderExists(targetSubfolder) Then
        fso.CreateFolder targetSubfolder
    End If

    If createWordBlank Then
        expectedCount = expectedCount + 1
        operationError = ""

        If wordApp Is Nothing Then
            Set wordApp = _
                NHblank_WordTemplate.NHblank_CreateWordApplication( _
                    operationError)
        End If

        If wordApp Is Nothing Then
            AddOperationError errorMsg, operationError
        Else
            fileName = NHblank_Utils.CreateFileNameWithPeriod( _
                lastName, firstName, discipline, CDate(referenceDate), "")
            wordOutputPath = targetSubfolder & "\" & fileName & ".docx"

            If NHblank_WordTemplate.NHblank_CreateFilledWordBlank( _
                wordApp, wordTemplatePath, wordOutputPath, _
                bgNumber, studentName, birthDateText, PROVIDER_NAME, _
                teacherName, discipline, overwriteExisting, _
                operationError) Then
                createdCount = createdCount + 1
            Else
                AddOperationError errorMsg, operationError
            End If
        End If
    End If

    ' Every Sozialamt/non-Jobcenter record receives the legacy Excel blank.
    If Not isJobcenter Then
        expectedCount = expectedCount + 1

        If Not IsDate(dateFrom) Or Not IsDate(dateTo) Then
            AddOperationError errorMsg, _
                "SA-Blank: Bewilligungszeitraum in G/H ist ungueltig."
        Else
            fileName = NHblank_Utils.CreateFileNameWithPeriod( _
                lastName, firstName, discipline, _
                CDate(referenceDate), "SA")
            legacyOutputPath = _
                targetSubfolder & "\" & fileName & ".xlsx"

            operationError = ""
            If CreateLegacyBlank( _
                legacyTemplatePath, legacyOutputPath, _
                lastName, firstName, discipline, _
                CDate(dateFrom), CDate(dateTo), _
                CDate(referenceDate), overwriteExisting, _
                operationError) Then
                createdCount = createdCount + 1
            Else
                AddOperationError errorMsg, operationError
            End If
        End If
    End If

    ProcessSingleRecord = (createdCount = expectedCount)
    Exit Function

ErrorHandler:
    AddOperationError errorMsg, _
        "Fehler #" & CStr(Err.Number) & ": " & Err.Description
    ProcessSingleRecord = False
End Function

Private Function CreateLegacyBlank( _
    ByVal templatePath As String, _
    ByVal outputPath As String, _
    ByVal lastName As String, _
    ByVal firstName As String, _
    ByVal discipline As String, _
    ByVal dateFrom As Date, _
    ByVal dateTo As Date, _
    ByVal referenceDate As Date, _
    ByVal overwriteExisting As Boolean, _
    ByRef errorMsg As String _
) As Boolean

    Dim wbOutput As Workbook
    Dim wsTemplate As Worksheet
    Dim fso As Object
    Dim localExcelPath As String
    Dim stagingPath As String
    Dim operationStage As String

    errorMsg = ""
    On Error GoTo ErrorHandler
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(templatePath) Then
        errorMsg = "SA-Template nicht gefunden: " & templatePath
        Exit Function
    End If

    If fso.FileExists(outputPath) And Not overwriteExisting Then
        errorMsg = "SA-Datei existiert bereits und wurde uebersprungen: " & _
                   outputPath
        Exit Function
    End If

    ' Excel can exhibit the same Unicode-path issue as Word when a workbook
    ' is opened from VBA. Work on a local TEMP copy and move the completed
    ' file into the teacher folder only after it has been closed.
    operationStage = "lokale SA-Arbeitsdatei vorbereiten"
    localExcelPath = CreateLocalLegacyWorkingPath(fso)
    fso.CopyFile templatePath, localExcelPath, True

    operationStage = "lokale SA-Arbeitsdatei in Excel oeffnen"
    Set wbOutput = Workbooks.Open( _
        FileName:=fso.GetFile(localExcelPath).ShortPath, _
        UpdateLinks:=0, _
        ReadOnly:=False, _
        AddToMru:=False)

    On Error Resume Next
    Set wsTemplate = wbOutput.Worksheets("Muster")
    On Error GoTo ErrorHandler

    If wsTemplate Is Nothing Then
        errorMsg = "Worksheet 'Muster' fehlt im SA-Template."
        GoTo CleanFail
    End If

    wsTemplate.Range("B1").Value = lastName & " " & firstName
    wsTemplate.Range("E1").Value = discipline
    wsTemplate.Range("B2").Value = _
        NHblank_Utils.FormatDateRange(dateFrom, dateTo)
    wsTemplate.Range("C4").Value = _
        NHblank_Utils.GetGermanMonth(Month(referenceDate))
    wsTemplate.Range("E4").Value = Year(referenceDate)

    wbOutput.Save
    wbOutput.Close SaveChanges:=False
    Set wbOutput = Nothing

    operationStage = "fertige SA-Datei in Zielordner kopieren"
    stagingPath = CreateTemporaryLegacyPath(fso, outputPath)
    fso.CopyFile localExcelPath, stagingPath, True

    operationStage = "SA-Zieldatei ersetzen"
    If fso.FileExists(outputPath) Then fso.DeleteFile outputPath, True
    fso.MoveFile stagingPath, outputPath
    fso.DeleteFile localExcelPath, True

    CreateLegacyBlank = True
    Exit Function

CleanFail:
    On Error Resume Next
    If Not wbOutput Is Nothing Then wbOutput.Close SaveChanges:=False
    Set wbOutput = Nothing
    If Not fso Is Nothing Then
        If Len(stagingPath) > 0 Then
            If fso.FileExists(stagingPath) Then fso.DeleteFile stagingPath, True
        End If
        If Len(localExcelPath) > 0 Then
            If fso.FileExists(localExcelPath) Then _
                fso.DeleteFile localExcelPath, True
        End If
    End If
    On Error GoTo 0
    CreateLegacyBlank = False
    Exit Function

ErrorHandler:
    errorMsg = "SA-Blank konnte nicht erstellt werden"
    If Len(operationStage) > 0 Then
        errorMsg = errorMsg & " (Schritt: " & operationStage & ")"
    End If
    errorMsg = errorMsg & ": " & Err.Description
    Resume CleanFail
End Function

Private Function CreateTemporaryLegacyPath( _
    ByVal fso As Object, _
    ByVal outputPath As String _
) As String

    Dim folderPath As String
    Dim candidate As String
    Dim counter As Long

    folderPath = Left$(outputPath, InStrRev(outputPath, "\") - 1)

    For counter = 1 To 1000
        candidate = folderPath & "\~nhblank_" & _
                    Format$(Now, "yyyymmdd_hhnnss") & "_" & _
                    Format$(counter, "000") & ".xlsx"
        If Not fso.FileExists(candidate) Then
            CreateTemporaryLegacyPath = candidate
            Exit Function
        End If
    Next counter

    Err.Raise vbObjectError + 2201, "NHblank_DataProcessor", _
              "Temporaerer SA-Dateiname konnte nicht erzeugt werden."
End Function

Private Function CreateLocalLegacyWorkingPath(ByVal fso As Object) As String
    Dim tempFolder As String
    Dim candidate As String
    Dim counter As Long

    tempFolder = fso.GetSpecialFolder(2).Path

    For counter = 1 To 1000
        candidate = fso.BuildPath( _
            tempFolder, _
            "nhblank_excel_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & _
            Format$(counter, "000") & ".xlsx")
        If Not fso.FileExists(candidate) Then
            CreateLocalLegacyWorkingPath = candidate
            Exit Function
        End If
    Next counter

    Err.Raise vbObjectError + 2202, "NHblank_DataProcessor", _
              "Lokale SA-Arbeitsdatei konnte nicht erzeugt werden."
End Function

Private Sub AddOperationError(ByRef errorList As String, _
                              ByVal message As String)
    If Len(Trim$(message)) = 0 Then Exit Sub

    If Len(errorList) > 0 Then errorList = errorList & vbCrLf
    errorList = errorList & message
End Sub

Private Sub AppendWarning(ByRef warningList As String, _
                          ByVal rowNum As Long, _
                          ByVal message As String)
    Const MAX_WARNING_TEXT As Long = 6000

    If Len(Trim$(message)) = 0 Then Exit Sub

    If Len(warningList) >= MAX_WARNING_TEXT Then Exit Sub

    If Len(warningList) > 0 Then warningList = warningList & vbCrLf & vbCrLf
    warningList = warningList & "Zeile " & CStr(rowNum) & ":" & vbCrLf & _
                  message

    If Len(warningList) >= MAX_WARNING_TEXT Then
        warningList = Left$(warningList, MAX_WARNING_TEXT) & vbCrLf & _
                      "... weitere Meldungen wurden gekuerzt."
    End If
End Sub

' Creates two synthetic files for deployment verification without reading
' personal data from Kinder_Blanks. The caller supplies a disposable folder.
Public Function NHblank_GenerationSelfTest( _
    ByVal wordTemplatePath As String, _
    ByVal legacyTemplatePath As String, _
    ByVal outputFolder As String _
) As String

    Dim wordApp As Object
    Dim wordPath As String
    Dim legacyPath As String
    Dim errorMsg As String

    On Error GoTo ErrorHandler

    If Len(Dir$(outputFolder, vbDirectory)) = 0 Then
        MkDir outputFolder
    End If

    wordPath = outputFolder & "\NHblank_Test_Sachunterricht_Juli_2026.docx"
    legacyPath = _
        outputFolder & "\NHblank_Test_Sachunterricht_SA_Juli_2026.xlsx"

    Set wordApp = _
        NHblank_WordTemplate.NHblank_CreateWordApplication(errorMsg)
    If wordApp Is Nothing Then
        NHblank_GenerationSelfTest = "FAILED: " & errorMsg
        Exit Function
    End If

    If Not NHblank_WordTemplate.NHblank_CreateFilledWordBlank( _
        wordApp, wordTemplatePath, wordPath, _
        "12345//1234567", "Mustermann Max", "15.03.2014", _
        PROVIDER_NAME, "Beispiel Anna", "Sachunterricht", _
        True, errorMsg) Then
        NHblank_GenerationSelfTest = "FAILED WORD: " & errorMsg
        GoTo CleanExit
    End If

    If Not CreateLegacyBlank( _
        legacyTemplatePath, legacyPath, _
        "Mustermann", "Max", "Sachunterricht", _
        DateSerial(2026, 7, 1), DateSerial(2026, 7, 31), _
        DateSerial(2026, 7, 1), True, errorMsg) Then
        NHblank_GenerationSelfTest = "FAILED LEGACY: " & errorMsg
        GoTo CleanExit
    End If

    NHblank_GenerationSelfTest = "OK"

CleanExit:
    On Error Resume Next
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordApp = Nothing
    On Error GoTo 0
    Exit Function

ErrorHandler:
    NHblank_GenerationSelfTest = "FAILED: " & Err.Description
    Resume CleanExit
End Function
