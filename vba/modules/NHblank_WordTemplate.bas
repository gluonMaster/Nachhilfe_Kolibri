Attribute VB_Name = "NHblank_WordTemplate"
Option Explicit

' =============================================================================
' NHblank_WordTemplate
' Creates editable Jobcenter attendance forms from Shablon.docx.
' Word is controlled through late binding; no Word reference is required.
' =============================================================================

Private Const ATTENDANCE_ROW_COUNT As Long = 17
Private Const EXPECTED_CONTROL_COUNT As Long = 144
Private Const EXPECTED_PAGE_COUNT As Long = 1
Private Const WD_STATISTIC_PAGES As Long = 2
Private Const WD_LINE_SPACE_SINGLE As Long = 0
Private Const FORM_FIELD_FONT_SIZE As Single = 10

Public Function NHblank_CreateWordApplication(ByRef errorMsg As String) As Object
    Dim wordApp As Object

    errorMsg = ""
    On Error GoTo ErrorHandler

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0

    Set NHblank_CreateWordApplication = wordApp
    Exit Function

ErrorHandler:
    errorMsg = "Microsoft Word konnte nicht gestartet werden: " & _
               Err.Description
    Set NHblank_CreateWordApplication = Nothing
End Function

Public Function NHblank_CheckWordTemplate( _
    ByVal templatePath As String, _
    ByRef errorMsg As String _
) As Boolean

    Dim wordApp As Object
    Dim doc As Object

    errorMsg = ""
    On Error GoTo ErrorHandler

    If Len(Dir$(templatePath)) = 0 Then
        errorMsg = "Word-Template nicht gefunden: " & templatePath
        Exit Function
    End If

    Set wordApp = NHblank_CreateWordApplication(errorMsg)
    If wordApp Is Nothing Then Exit Function

    ' FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
    Set doc = wordApp.Documents.Open(templatePath, False, True, False)

    If Not ValidateDocumentStructure(doc, errorMsg) Then GoTo CleanFail

    If doc.ComputeStatistics(WD_STATISTIC_PAGES) <> EXPECTED_PAGE_COUNT Then
        errorMsg = "Word-Template hat nicht genau eine Seite."
        GoTo CleanFail
    End If

    doc.Close False
    wordApp.Quit
    Set doc = Nothing
    Set wordApp = Nothing

    NHblank_CheckWordTemplate = True
    Exit Function

CleanFail:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    Set doc = Nothing
    Set wordApp = Nothing
    NHblank_CheckWordTemplate = False
    Exit Function

ErrorHandler:
    errorMsg = "Word-Template konnte nicht geprueft werden: " & _
               Err.Description
    Resume CleanFail
End Function

Public Function NHblank_CreateFilledWordBlank( _
    ByVal wordApp As Object, _
    ByVal templatePath As String, _
    ByVal outputPath As String, _
    ByVal bgNumber As String, _
    ByVal studentName As String, _
    ByVal birthDateText As String, _
    ByVal providerName As String, _
    ByVal teacherName As String, _
    ByVal discipline As String, _
    ByVal overwriteExisting As Boolean, _
    ByRef errorMsg As String _
) As Boolean

    Dim doc As Object
    Dim fso As Object
    Dim localWordPath As String
    Dim wordOpenPath As String
    Dim stagingPath As String
    Dim operationStage As String
    Dim shortenedDiscipline As String
    Dim rowNumber As Long

    errorMsg = ""
    On Error GoTo ErrorHandler
    Set fso = CreateObject("Scripting.FileSystemObject")

    If wordApp Is Nothing Then
        errorMsg = "Microsoft Word ist nicht initialisiert."
        Exit Function
    End If

    If Not fso.FileExists(templatePath) Then
        errorMsg = "Word-Template nicht gefunden: " & templatePath
        Exit Function
    End If

    If fso.FileExists(outputPath) And Not overwriteExisting Then
        errorMsg = "Datei existiert bereits und wurde uebersprungen: " & _
                   outputPath
        Exit Function
    End If

    If Not ValidateTextLengths( _
        bgNumber, studentName, providerName, teacherName, errorMsg) Then
        Exit Function
    End If

    shortenedDiscipline = NHblank_AbbreviateDiscipline(discipline)
    If Len(shortenedDiscipline) > 15 Then
        errorMsg = "Fach ist auch nach der Abkuerzung laenger als " & _
                   "15 Zeichen: " & shortenedDiscipline
        Exit Function
    End If

    ' Word opened from Excel/VBA can fail on a path containing characters
    ' outside the active Windows code page (for example a teacher-name umlaut).
    ' document that Word itself opens in the local TEMP folder. Only after the
    ' document is complete do we copy it into the Unicode target folder.
    operationStage = "lokale Arbeitsdatei vorbereiten"
    localWordPath = CreateLocalWordWorkingPath(fso, ".docx")
    fso.CopyFile templatePath, localWordPath, True
    wordOpenPath = fso.GetFile(localWordPath).ShortPath
    If Len(wordOpenPath) = 0 Then wordOpenPath = localWordPath

    ' FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
    operationStage = "lokale Arbeitsdatei in Word oeffnen"
    Set doc = wordApp.Documents.Open(wordOpenPath, False, False, False)

    If Not ValidateDocumentStructure(doc, errorMsg) Then GoTo CleanFail

    SetFixedSizeControlText doc, "BG_NR", bgNumber
    SetFixedSizeControlText doc, "SCHUELER_NAME", studentName
    SetFixedSizeControlText doc, "GEBURTSDATUM", birthDateText
    SetControlText doc, "ANBIETER_NAME_FIRMA", providerName
    SetFixedSizeControlText doc, "LEHRKRAFT_NAME", teacherName

    For rowNumber = 1 To ATTENDANCE_ROW_COUNT
        SetControlText doc, _
                       AttendanceTag(rowNumber, "FACH"), _
                       shortenedDiscipline
    Next rowNumber

    doc.Repaginate
    If doc.ComputeStatistics(WD_STATISTIC_PAGES) <> EXPECTED_PAGE_COUNT Then
        errorMsg = "Der ausgefuellte Word-Blank hat mehr als eine Seite."
        GoTo CleanFail
    End If

    doc.Save
    doc.Close False
    Set doc = Nothing

    operationStage = "fertiges Dokument in Zielordner kopieren"
    stagingPath = CreateTemporarySiblingPath(fso, outputPath, ".docx")
    fso.CopyFile localWordPath, stagingPath, True

    operationStage = "Zieldatei ersetzen"
    If fso.FileExists(outputPath) Then fso.DeleteFile outputPath, True
    fso.MoveFile stagingPath, outputPath
    fso.DeleteFile localWordPath, True

    NHblank_CreateFilledWordBlank = True
    Exit Function

CleanFail:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    Set doc = Nothing
    If Not fso Is Nothing Then
        If Len(stagingPath) > 0 Then
            If fso.FileExists(stagingPath) Then fso.DeleteFile stagingPath, True
        End If
        If Len(localWordPath) > 0 Then
            If fso.FileExists(localWordPath) Then _
                fso.DeleteFile localWordPath, True
        End If
    End If
    On Error GoTo 0
    NHblank_CreateFilledWordBlank = False
    Exit Function

ErrorHandler:
    errorMsg = "Word-Blank konnte nicht erstellt werden"
    If Len(operationStage) > 0 Then
        errorMsg = errorMsg & " (Schritt: " & operationStage & ")"
    End If
    errorMsg = errorMsg & ": " & Err.Description
    Resume CleanFail
End Function

Public Function NHblank_AbbreviateDiscipline( _
    ByVal discipline As String _
) As String

    Dim cleaned As String

    cleaned = CleanControlText(discipline)

    Select Case LCase$(cleaned)
        Case "sachunterricht"
            NHblank_AbbreviateDiscipline = "Sachunt."
        Case Else
            NHblank_AbbreviateDiscipline = cleaned
    End Select
End Function

Private Function ValidateDocumentStructure( _
    ByVal doc As Object, _
    ByRef errorMsg As String _
) As Boolean

    Dim rowNumber As Long
    Dim suffix As Variant
    Dim suffixes As Variant

    If doc.ContentControls.Count <> EXPECTED_CONTROL_COUNT Then
        errorMsg = "Unerwartete Anzahl Word-Felder: " & _
                   CStr(doc.ContentControls.Count) & _
                   "; erwartet: " & CStr(EXPECTED_CONTROL_COUNT)
        Exit Function
    End If

    If Not AssertSingleTag(doc, "BG_NR", errorMsg) Then Exit Function
    If Not AssertSingleTag(doc, "SCHUELER_NAME", errorMsg) Then Exit Function
    If Not AssertSingleTag(doc, "GEBURTSDATUM", errorMsg) Then Exit Function
    If Not AssertSingleTag( _
        doc, "UNTERSCHRIFT_ERZIEHUNGSBERECHTIGTE", errorMsg) Then Exit Function
    If Not AssertSingleTag( _
        doc, "ANBIETER_NAME_FIRMA", errorMsg) Then Exit Function
    If Not AssertSingleTag( _
        doc, "ANBIETER_STEMPEL_UNTERSCHRIFT_DATUM", errorMsg) Then Exit Function
    If Not AssertSingleTag(doc, "LEHRKRAFT_NAME", errorMsg) Then Exit Function
    If Not AssertSingleTag( _
        doc, "LEHRKRAFT_UNTERSCHRIFT", errorMsg) Then Exit Function

    suffixes = Array( _
        "DATUM", _
        "UHRZEIT", _
        "DAUER", _
        "GR_EZ", _
        "FACH", _
        "ORT", _
        "STATUS", _
        "ABWESENHEITSGRUND")

    For rowNumber = 1 To ATTENDANCE_ROW_COUNT
        For Each suffix In suffixes
            If Not AssertSingleTag( _
                doc, AttendanceTag(rowNumber, CStr(suffix)), errorMsg) Then
                Exit Function
            End If
        Next suffix
    Next rowNumber

    ValidateDocumentStructure = True
End Function

Private Function AssertSingleTag(ByVal doc As Object, _
                                 ByVal tagName As String, _
                                 ByRef errorMsg As String) As Boolean
    Dim controls As Object

    Set controls = doc.SelectContentControlsByTag(tagName)

    If controls.Count <> 1 Then
        errorMsg = "Word-Feld muss genau einmal vorhanden sein: " & tagName
        Exit Function
    End If

    AssertSingleTag = True
End Function

Private Sub SetControlText(ByVal doc As Object, _
                           ByVal tagName As String, _
                           ByVal textValue As String)
    Dim controls As Object
    Dim control As Object
    Dim safeText As String

    Set controls = doc.SelectContentControlsByTag(tagName)
    If controls.Count <> 1 Then
        Err.Raise vbObjectError + 2101, "NHblank_WordTemplate", _
                  "Word-Feld nicht eindeutig: " & tagName
    End If

    safeText = CleanControlText(textValue)
    If Len(safeText) = 0 Then safeText = " "

    Set control = controls.Item(1)
    control.LockContents = False
    control.LockContentControl = False
    control.Range.Text = safeText
End Sub

Private Sub SetFixedSizeControlText(ByVal doc As Object, _
                                    ByVal tagName As String, _
                                    ByVal textValue As String)
    Dim controls As Object
    Dim control As Object

    SetControlText doc, tagName, textValue

    Set controls = doc.SelectContentControlsByTag(tagName)
    Set control = controls.Item(1)

    With control.Range
        .Font.Size = FORM_FIELD_FONT_SIZE
        .Font.Position = 0

        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = WD_LINE_SPACE_SINGLE
        End With
    End With
End Sub

Private Function AttendanceTag(ByVal rowNumber As Long, _
                               ByVal suffix As String) As String
    If rowNumber < 1 Or rowNumber > ATTENDANCE_ROW_COUNT Then
        Err.Raise vbObjectError + 2102, "NHblank_WordTemplate", _
                  "Ungueltige Anwesenheitszeile: " & CStr(rowNumber)
    End If

    AttendanceTag = "R" & Format$(rowNumber, "00") & "_" & suffix
End Function

Private Function CleanControlText(ByVal textValue As String) As String
    Dim result As String

    result = Trim$(textValue)
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, vbTab, " ")

    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    CleanControlText = result
End Function

Private Function ValidateTextLengths( _
    ByVal bgNumber As String, _
    ByVal studentName As String, _
    ByVal providerName As String, _
    ByVal teacherName As String, _
    ByRef errorMsg As String _
) As Boolean

    If Len(CleanControlText(bgNumber)) > 30 Then
        errorMsg = "BG-Nummer ist laenger als 30 Zeichen."
        Exit Function
    End If

    If Len(CleanControlText(studentName)) > 40 Then
        errorMsg = "Schuelername ist laenger als 40 Zeichen."
        Exit Function
    End If

    If Len(CleanControlText(providerName)) > 40 Then
        errorMsg = "Name / Firma ist laenger als 40 Zeichen."
        Exit Function
    End If

    If Len(CleanControlText(teacherName)) > 35 Then
        errorMsg = "Lehrkraftname ist laenger als 35 Zeichen."
        Exit Function
    End If

    ValidateTextLengths = True
End Function

Private Function CreateTemporarySiblingPath( _
    ByVal fso As Object, _
    ByVal outputPath As String, _
    ByVal extensionWithDot As String _
) As String

    Dim folderPath As String
    Dim candidate As String
    Dim counter As Long

    folderPath = Left$(outputPath, InStrRev(outputPath, "\") - 1)

    For counter = 1 To 1000
        candidate = folderPath & "\~nhblank_" & _
                    Format$(Now, "yyyymmdd_hhnnss") & "_" & _
                    Format$(counter, "000") & extensionWithDot
        If Not fso.FileExists(candidate) Then
            CreateTemporarySiblingPath = candidate
            Exit Function
        End If
    Next counter

    Err.Raise vbObjectError + 2103, "NHblank_WordTemplate", _
              "Temporaerer Dateiname konnte nicht erzeugt werden."
End Function

Private Function CreateLocalWordWorkingPath( _
    ByVal fso As Object, _
    ByVal extensionWithDot As String _
) As String

    Dim tempFolder As String
    Dim candidate As String
    Dim counter As Long

    tempFolder = fso.GetSpecialFolder(2).Path

    For counter = 1 To 1000
        candidate = fso.BuildPath( _
            tempFolder, _
            "nhblank_word_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & _
            Format$(counter, "000") & extensionWithDot)
        If Not fso.FileExists(candidate) Then
            CreateLocalWordWorkingPath = candidate
            Exit Function
        End If
    Next counter

    Err.Raise vbObjectError + 2104, "NHblank_WordTemplate", _
              "Lokale Word-Arbeitsdatei konnte nicht erzeugt werden."
End Function
