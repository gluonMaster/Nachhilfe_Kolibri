Attribute VB_Name = "NHblank_AdminSync"
Option Explicit

' =============================================================================
' NHblank_AdminSync
' Full operator workflow:
'   1. Remove Kinder rows cancelled in Admin Kartei with status KN.
'   2. Refresh teacher (I) and end date (H) on Kinder from Admin Kartei.
'   3. Hard-sync active Kinder rows into Kinder_Blanks with a stable key.
' =============================================================================

Private Const KARTEI_FIRST_DATA_ROW As Long = 2
Private Const KINDER_FIRST_DATA_ROW As Long = 5
Private Const BLANKS_FIRST_DATA_ROW As Long = 5

Private Const COL_A As String = "A"
Private Const COL_B As String = "B"
Private Const COL_C As String = "C"
Private Const COL_D As String = "D"
Private Const COL_E As String = "E"
Private Const COL_F As String = "F"
Private Const COL_G As String = "G"
Private Const COL_H As String = "H"
Private Const COL_I As String = "I"
Private Const COL_K As String = "K"
Private Const COL_L As String = "L"
Private Const COL_O As String = "O"
Private Const COL_P As String = "P"
Private Const COL_S As String = "S"
Private Const COL_T As String = "T"

Private Type AdminSyncStats
    KinderChecked As Long
    KinderDeletedKN As Long
    TeacherUpdated As Long
    DateUpdated As Long
    BlanksAdded As Long
    BlanksUpdated As Long
    BlanksDeleted As Long
    DuplicateSkipped As Long
    WarningCount As Long
End Type

' -----------------------------------------------------------------------------
' Public implementation entry point.
' Called by NHblank_Menu.NHblank_RunFullAdminSync.
' -----------------------------------------------------------------------------
Public Sub NHblank_AdminSync_Run()
    Dim appState As NHblank_AppState
    Dim stats As AdminSyncStats
    Dim wbAdmin As Workbook
    Dim wsKartei As Worksheet
    Dim wsKinder As Worksheet
    Dim wsErrorLog As Worksheet
    Dim adminIndex As Object
    Dim monthNumber As Long
    Dim semester As Long
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo ErrorHandler

    Set appState = New NHblank_AppState
    appState.Capture
    appState.OptimizeForRun

    If InStr(1, ThisWorkbook.name, "Nachhilfe", vbTextCompare) = 0 Then
        MsgBox "Bitte fuehren Sie dieses Makro aus der Nachhilfe-Datei aus.", _
               vbExclamation, "Falsche Datei"
        GoTo CleanExit
    End If

    If Not Kind_TryGetOpenWorkbookByName(KIND_ADMIN_WB_NAME, wbAdmin) Then
        MsgBox "Die Admin-Datei '" & KIND_ADMIN_WB_NAME & "' ist nicht geoeffnet." & vbCrLf & _
               "Bitte oeffnen Sie die Datei und starten Sie die Synchronisation erneut.", _
               vbExclamation, "Admin-Datei nicht gefunden"
        GoTo CleanExit
    End If

    If Not Kind_TryGetWorksheet(wbAdmin, KIND_ADMIN_SHEET_KARTEI, wsKartei) Then
        MsgBox "Das Blatt '" & KIND_ADMIN_SHEET_KARTEI & "' wurde in der Admin-Datei nicht gefunden.", _
               vbCritical, "Blatt nicht gefunden"
        GoTo CleanExit
    End If

    If Not Kind_TryGetWorksheet(ThisWorkbook, KIND_THIS_SHEET_KINDER, wsKinder) Then
        MsgBox "Das Blatt '" & KIND_THIS_SHEET_KINDER & "' wurde in der Nachhilfe-Datei nicht gefunden.", _
               vbCritical, "Blatt nicht gefunden"
        GoTo CleanExit
    End If

    Set wsErrorLog = EnsureErrorLogSheet()
    If wsErrorLog Is Nothing Then
        MsgBox "Das Blatt 'ErrorLog' konnte nicht erstellt werden.", vbCritical, "Log-Fehler"
        GoTo CleanExit
    End If

    If Not DetermineSemesterFromKinder(wsKinder, monthNumber, semester) Then
        MsgBox "Ungueltiges Referenzdatum in Kinder!T2.", vbExclamation, "Datum fehlt"
        GoTo CleanExit
    End If

    NHblank_ApplyAllFormats wsKinder, KINDER_FIRST_DATA_ROW

    Application.StatusBar = "Admin-Index wird aufgebaut..."
    Set adminIndex = BuildAdminIndex(wsKartei, semester)
    If adminIndex Is Nothing Then
        MsgBox "Der Admin-Index konnte nicht aufgebaut werden.", vbCritical, "Index-Fehler"
        GoTo CleanExit
    End If

    Application.StatusBar = "Kinder wird mit Admin synchronisiert..."
    SyncKinderWithAdmin wsKinder, adminIndex, stats

    Application.StatusBar = "Kinder_Blanks wird aktualisiert..."
    SyncKinderToBlanksStable wsKinder, stats

    ShowSummary stats, monthNumber, semester

CleanExit:
    If Not appState Is Nothing Then appState.Restore
    Exit Sub

ErrorHandler:
    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next
    If Not wsErrorLog Is Nothing Then
        LogIssue stats, "", "", "", "", "RUNTIME_ERR", _
                 "Fehler " & errNumber & ": " & errDescription
    End If
    If Not appState Is Nothing Then appState.Restore
    On Error GoTo 0

    MsgBox "Fehler bei der Admin-Synchronisation:" & vbCrLf & _
           errDescription, vbCritical, "Fehler"
End Sub

' =============================================================================
' Admin -> Kinder
' =============================================================================

Private Function BuildAdminIndex(ByVal wsKartei As Worksheet, ByVal semester As Long) As Object
    Dim dict As Object
    Dim lastRow As Long
    Dim rowNum As Long
    Dim subjectCol As String
    Dim teacherCol As String
    Dim familyIDNorm As String
    Dim subjectNorm As String
    Dim subjectRaw As String
    Dim dictKey As String
    Dim coll As Collection
    Dim entry As Variant

    On Error GoTo ErrHandler

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    subjectCol = GetSubjectColumnForSemester(semester)
    teacherCol = GetTeacherColumnForSemester(semester)

    lastRow = wsKartei.Cells(wsKartei.rows.Count, KIND_KARTEI_COL_FAMILYID).End(xlUp).row

    For rowNum = KARTEI_FIRST_DATA_ROW To lastRow
        familyIDNorm = Kind_NormalizeForCompare(ValueTextOrEmpty(wsKartei.Range(KIND_KARTEI_COL_FAMILYID & rowNum).value))
        If Len(familyIDNorm) = 0 Then GoTo NextRow

        subjectRaw = ValueTextOrEmpty(wsKartei.Range(subjectCol & rowNum).value)
        If Not Kind_TryNormalizeSubjectKartei(subjectRaw, subjectNorm) Then GoTo NextRow

        dictKey = familyIDNorm & "|" & subjectNorm
        entry = Array( _
            rowNum, _
            ValueTextOrEmpty(wsKartei.Range(KIND_KARTEI_COL_CHILD & rowNum).value), _
            ValueTextOrEmpty(wsKartei.Range(KIND_KARTEI_COL_STATUS & rowNum).value), _
            ValueTextOrEmpty(wsKartei.Range(teacherCol & rowNum).value), _
            wsKartei.Range(COL_C & rowNum).value _
        )

        If dict.exists(dictKey) Then
            Set coll = dict(dictKey)
        Else
            Set coll = New Collection
            dict.Add dictKey, coll
        End If
        coll.Add entry

NextRow:
    Next rowNum

    Set BuildAdminIndex = dict
    Exit Function

ErrHandler:
    Set BuildAdminIndex = Nothing
End Function

Private Sub SyncKinderWithAdmin(ByVal wsKinder As Worksheet, _
                                ByVal adminIndex As Object, _
                                ByRef stats As AdminSyncStats)
    Dim lastRow As Long
    Dim rowNum As Long
    Dim matchFound As Boolean
    Dim adminRow As Long
    Dim adminStatus As String
    Dim adminTeacher As String
    Dim adminDateRaw As Variant
    Dim reasonCode As String
    Dim details As String
    Dim parsedDate As Date
    Dim familyID As String
    Dim lastName As String
    Dim firstName As String
    Dim subject As String

    lastRow = GetLastDataRow(wsKinder)
    If lastRow < KINDER_FIRST_DATA_ROW Then Exit Sub

    For rowNum = lastRow To KINDER_FIRST_DATA_ROW Step -1
        If Not RowHasKinderIdentity(wsKinder, rowNum) Then GoTo NextRow

        familyID = ValueTextOrEmpty(wsKinder.Range(COL_B & rowNum).value)
        lastName = ValueTextOrEmpty(wsKinder.Range(COL_C & rowNum).value)
        firstName = ValueTextOrEmpty(wsKinder.Range(COL_D & rowNum).value)
        subject = ValueTextOrEmpty(wsKinder.Range(COL_E & rowNum).value)

        stats.KinderChecked = stats.KinderChecked + 1

        matchFound = FindAdminMatch(adminIndex, familyID, lastName, firstName, subject, _
                                    adminRow, adminStatus, adminTeacher, adminDateRaw, _
                                    reasonCode, details)

        If Not matchFound Then
            LogIssue stats, familyID, lastName, firstName, subject, reasonCode, details
            GoTo NextRow
        End If

        If UCase$(Trim$(adminStatus)) = "KN" Then
            wsKinder.rows(rowNum).Delete
            stats.KinderDeletedKN = stats.KinderDeletedKN + 1
            GoTo NextRow
        End If

        adminTeacher = Trim$(adminTeacher)
        If Len(adminTeacher) > 0 Then
            If StrComp(ValueTextOrEmpty(wsKinder.Range(COL_I & rowNum).value), adminTeacher, vbTextCompare) <> 0 Then
                wsKinder.Range(COL_I & rowNum).value = adminTeacher
                stats.TeacherUpdated = stats.TeacherUpdated + 1
            End If
        End If

        If TryParseKarteiDate(adminDateRaw, parsedDate) Then
            If Not CellDateEquals(wsKinder.Range(COL_H & rowNum).value, parsedDate) Then
                wsKinder.Range(COL_H & rowNum).value = parsedDate
                wsKinder.Range(COL_H & rowNum).NumberFormat = "dd.mm.yyyy"
                stats.DateUpdated = stats.DateUpdated + 1
            End If
        Else
            If IsBlankValue(adminDateRaw) Then
                LogIssue stats, familyID, lastName, firstName, subject, "DATE_EMPTY", _
                         "Admin Kartei Zeile " & adminRow & ": Datum in Spalte C ist leer"
            Else
                LogIssue stats, familyID, lastName, firstName, subject, "DATE_INVALID", _
                         "Admin Kartei Zeile " & adminRow & ": Datum nicht erkannt: " & SafeValueText(adminDateRaw)
            End If
        End If

NextRow:
    Next rowNum
End Sub

Private Function FindAdminMatch(ByVal adminIndex As Object, _
                                ByVal familyID As String, _
                                ByVal kinderLast As String, _
                                ByVal kinderFirst As String, _
                                ByVal kinderSubject As String, _
                                ByRef outAdminRow As Long, _
                                ByRef outStatus As String, _
                                ByRef outTeacher As String, _
                                ByRef outDateRaw As Variant, _
                                ByRef outReasonCode As String, _
                                ByRef outDetails As String) As Boolean
    Dim dictKey As String
    Dim subjectNorm As String
    Dim familyIDNorm As String
    Dim coll As Collection
    Dim entry As Variant
    Dim matchCount As Long
    Dim matchingRows As String
    Dim i As Long

    FindAdminMatch = False
    outAdminRow = 0
    outStatus = ""
    outTeacher = ""
    outDateRaw = Empty
    outReasonCode = ""
    outDetails = ""

    If adminIndex Is Nothing Then
        outReasonCode = "NO_INDEX"
        outDetails = "Admin-Index nicht verfuegbar"
        Exit Function
    End If

    familyIDNorm = Kind_NormalizeForCompare(familyID)
    subjectNorm = Kind_NormalizeSubjectKinder(kinderSubject)
    dictKey = familyIDNorm & "|" & subjectNorm

    If Not adminIndex.exists(dictKey) Then
        outReasonCode = "NO_KEY"
        outDetails = "Kein Admin-Treffer fuer FamilyID/Fach"
        Exit Function
    End If

    Set coll = adminIndex(dictKey)

    For i = 1 To coll.Count
        entry = coll(i)
        If Kind_IsSameChild(CStr(entry(1)), kinderLast, kinderFirst) Then
            matchCount = matchCount + 1
            If Len(matchingRows) > 0 Then matchingRows = matchingRows & ", "
            matchingRows = matchingRows & CStr(entry(0))

            If matchCount = 1 Then
                outAdminRow = CLng(entry(0))
                outStatus = CStr(entry(2))
                outTeacher = CStr(entry(3))
                outDateRaw = entry(4)
            End If
        End If
    Next i

    Select Case matchCount
        Case 0
            outReasonCode = "NO_CHILD"
            outDetails = "FamilyID/Fach gefunden, aber Kind-Name passt nicht (" & coll.Count & " Admin-Zeilen)"
        Case 1
            FindAdminMatch = True
        Case Else
            outReasonCode = "AMBIGUOUS"
            outDetails = "Mehrere Admin-Treffer in Zeilen: " & matchingRows
    End Select
End Function

' =============================================================================
' Kinder -> Kinder_Blanks full hard sync
' =============================================================================

Private Sub SyncKinderToBlanksStable(ByVal wsKinder As Worksheet, ByRef stats As AdminSyncStats)
    Dim wsBlanks As Worksheet
    Dim dictKinder As Object
    Dim dictBlanks As Object
    Dim lastRowKinder As Long
    Dim lastRowBlanks As Long
    Dim key As Variant
    Dim kinderRows As Collection
    Dim blanksRows As Collection
    Dim kinderRow As Long
    Dim blanksRow As Long
    Dim rowsToDelete As Collection

    Set wsBlanks = NHblank_EnsureKinderBlanksSheetExists()
    If wsBlanks Is Nothing Then
        LogIssue stats, "", "", "", "", "BLANKS_MISSING", _
                 "Kinder_Blanks konnte nicht erstellt oder gefunden werden"
        Exit Sub
    End If

    NHblank_ApplyAllFormats wsKinder, KINDER_FIRST_DATA_ROW
    NHblank_ApplyAllFormats wsBlanks, BLANKS_FIRST_DATA_ROW

    wsBlanks.Range("T2").value = wsKinder.Range("T2").value
    wsBlanks.Range("T2").NumberFormat = "dd.mm.yyyy"

    Set dictKinder = CreateObject("Scripting.Dictionary")
    Set dictBlanks = CreateObject("Scripting.Dictionary")
    dictKinder.CompareMode = vbTextCompare
    dictBlanks.CompareMode = vbTextCompare

    lastRowKinder = GetLastDataRow(wsKinder)
    lastRowBlanks = GetLastDataRow(wsBlanks)

    If lastRowKinder >= KINDER_FIRST_DATA_ROW Then
        BuildStableKeyDictionary wsKinder, KINDER_FIRST_DATA_ROW, lastRowKinder, dictKinder, True
    End If

    If lastRowBlanks >= BLANKS_FIRST_DATA_ROW Then
        BuildStableKeyDictionary wsBlanks, BLANKS_FIRST_DATA_ROW, lastRowBlanks, dictBlanks, False
    End If

    For Each key In dictKinder.Keys
        Set kinderRows = dictKinder(key)
        If kinderRows.Count > 1 Then
            LogDuplicateKey stats, wsKinder.name, CStr(key), kinderRows
            GoTo NextKinderKey
        End If

        kinderRow = CLng(kinderRows(1))

        If dictBlanks.exists(key) Then
            Set blanksRows = dictBlanks(key)
            If blanksRows.Count > 1 Then
                LogDuplicateKey stats, wsBlanks.name, CStr(key), blanksRows
                GoTo NextKinderKey
            End If

            blanksRow = CLng(blanksRows(1))
            If UpdateBlanksRow(wsKinder, kinderRow, wsBlanks, blanksRow) Then
                stats.BlanksUpdated = stats.BlanksUpdated + 1
            End If
        Else
            AddRecordToBlanksStable wsKinder, kinderRow, wsBlanks
            stats.BlanksAdded = stats.BlanksAdded + 1
        End If

NextKinderKey:
    Next key

    Set rowsToDelete = New Collection
    For Each key In dictBlanks.Keys
        Set blanksRows = dictBlanks(key)

        If Not dictKinder.exists(key) Then
            AddRowsToCollection blanksRows, rowsToDelete
        Else
            Set kinderRows = dictKinder(key)
            If kinderRows.Count > 1 Then
                ' Source duplicate is ambiguous. Keep existing blanks rows untouched.
            ElseIf blanksRows.Count > 1 Then
                ' Target duplicate is ambiguous while source exists. Keep untouched.
            End If
        End If
    Next key

    stats.BlanksDeleted = stats.BlanksDeleted + DeleteRowsFromBottomUp(wsBlanks, rowsToDelete)

    SortByColumnC wsBlanks
    RenumberColumnA wsBlanks
    NHblank_ApplyAllFormats wsBlanks, BLANKS_FIRST_DATA_ROW
End Sub

Private Sub BuildStableKeyDictionary(ByVal ws As Worksheet, _
                                     ByVal startRow As Long, _
                                     ByVal endRow As Long, _
                                     ByRef dict As Object, _
                                     ByVal activeOnly As Boolean)
    Dim rowNum As Long
    Dim key As String
    Dim rowsForKey As Collection

    For rowNum = startRow To endRow
        If Len(ValueTextOrEmpty(ws.Range(COL_C & rowNum).value)) = 0 Then GoTo NextRow
        If activeOnly Then
            If Not IsActiveRecordSafe(ws, rowNum) Then GoTo NextRow
        End If

        key = BuildStableRecordKey(ws, rowNum)
        If Len(Replace(key, "|", "")) = 0 Then GoTo NextRow

        If dict.exists(key) Then
            Set rowsForKey = dict(key)
        Else
            Set rowsForKey = New Collection
            dict.Add key, rowsForKey
        End If
        rowsForKey.Add rowNum

NextRow:
    Next rowNum
End Sub

Private Function BuildStableRecordKey(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    Dim keyParts(1 To 8) As String

    keyParts(1) = NormalizeKeyCell(ws.Range(COL_B & rowNum).value)
    keyParts(2) = NormalizeKeyCell(ws.Range(COL_C & rowNum).value)
    keyParts(3) = NormalizeKeyCell(ws.Range(COL_D & rowNum).value)
    keyParts(4) = NormalizeKeyCell(ws.Range(COL_E & rowNum).value)
    keyParts(5) = NormalizeKeyCell(ws.Range(COL_F & rowNum).value)
    keyParts(6) = NormalizeKeyCell(ws.Range(COL_K & rowNum).value)
    keyParts(7) = NormalizeKeyCell(ws.Range(COL_L & rowNum).value)
    keyParts(8) = NormalizeKeyCell(NHblank_BgNummerToString(ws.Range(COL_O & rowNum).value))

    BuildStableRecordKey = LCase$(Join(keyParts, "|"))
End Function

Private Function UpdateBlanksRow(ByVal wsKinder As Worksheet, _
                                 ByVal kinderRow As Long, _
                                 ByVal wsBlanks As Worksheet, _
                                 ByVal blanksRow As Long) As Boolean
    Dim changed As Boolean

    If SetCellIfDifferent(wsBlanks.Range(COL_G & blanksRow), wsKinder.Range(COL_G & kinderRow).value) Then changed = True
    If SetCellIfDifferent(wsBlanks.Range(COL_H & blanksRow), wsKinder.Range(COL_H & kinderRow).value) Then changed = True
    If SetCellIfDifferent(wsBlanks.Range(COL_I & blanksRow), wsKinder.Range(COL_I & kinderRow).value) Then changed = True
    If SetCellIfDifferent(wsBlanks.Range(COL_S & blanksRow), wsKinder.Range(COL_S & kinderRow).value) Then changed = True
    If SetCellIfDifferent(wsBlanks.Range(COL_T & blanksRow), wsKinder.Range(COL_T & kinderRow).value) Then changed = True

    wsBlanks.Range(COL_G & blanksRow).NumberFormat = "dd.mm.yyyy"
    wsBlanks.Range(COL_H & blanksRow).NumberFormat = "dd.mm.yyyy"

    UpdateBlanksRow = changed
End Function

Private Sub AddRecordToBlanksStable(ByVal wsKinder As Worksheet, _
                                    ByVal kinderRow As Long, _
                                    ByVal wsBlanks As Worksheet)
    Dim newRow As Long
    Dim colNum As Long

    newRow = GetLastDataRow(wsBlanks) + 1
    If newRow < BLANKS_FIRST_DATA_ROW Then newRow = BLANKS_FIRST_DATA_ROW

    For colNum = 2 To 20 ' B:T
        If colNum = 15 Then
            NHblank_WriteBgNummer wsBlanks.Cells(newRow, colNum), wsKinder.Cells(kinderRow, colNum).value
        Else
            wsBlanks.Cells(newRow, colNum).value = wsKinder.Cells(kinderRow, colNum).value
        End If
    Next colNum

    wsBlanks.Range(COL_G & newRow).NumberFormat = "dd.mm.yyyy"
    wsBlanks.Range(COL_H & newRow).NumberFormat = "dd.mm.yyyy"
End Sub

' =============================================================================
' Date parsing
' =============================================================================

Private Function TryParseKarteiDate(ByVal rawValue As Variant, ByRef outDate As Date) As Boolean
    Dim textValue As String
    Dim serialValue As Double
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim dayPart As Long
    Dim monthPart As Long
    Dim yearPart As Long

    TryParseKarteiDate = False

    If IsError(rawValue) Then Exit Function
    If IsNull(rawValue) Then Exit Function
    If IsEmpty(rawValue) Then Exit Function

    If VarType(rawValue) = vbDate Then
        outDate = dateValue(CDate(rawValue))
        TryParseKarteiDate = DateWithinReasonableRange(outDate)
        Exit Function
    End If

    If IsNumeric(rawValue) And VarType(rawValue) <> vbString Then
        serialValue = CDbl(rawValue)
        If serialValue >= CDbl(DateSerial(2000, 1, 1)) And _
           serialValue <= CDbl(DateSerial(2100, 12, 31)) Then
            outDate = dateValue(CDate(serialValue))
            TryParseKarteiDate = True
        End If
        Exit Function
    End If

    textValue = CleanDateText(CStr(rawValue))
    If Len(textValue) = 0 Then Exit Function

    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False
        .ignoreCase = True
        .Pattern = "(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})"
    End With

    If Not regex.Test(textValue) Then Exit Function

    Set matches = regex.Execute(textValue)
    Set match = matches(0)

    dayPart = CLng(match.SubMatches(0))
    monthPart = CLng(match.SubMatches(1))
    yearPart = CLng(match.SubMatches(2))

    If yearPart < 100 Then
        If yearPart < 50 Then
            yearPart = 2000 + yearPart
        Else
            yearPart = 1900 + yearPart
        End If
    End If

    If TryBuildDate(dayPart, monthPart, yearPart, outDate) Then
        TryParseKarteiDate = DateWithinReasonableRange(outDate)
    End If
End Function

Private Function TryBuildDate(ByVal dayPart As Long, _
                              ByVal monthPart As Long, _
                              ByVal yearPart As Long, _
                              ByRef outDate As Date) As Boolean
    Dim candidate As Date

    On Error GoTo ErrHandler

    candidate = DateSerial(yearPart, monthPart, dayPart)
    If Day(candidate) <> dayPart Then Exit Function
    If Month(candidate) <> monthPart Then Exit Function
    If Year(candidate) <> yearPart Then Exit Function

    outDate = candidate
    TryBuildDate = True
    Exit Function

ErrHandler:
    TryBuildDate = False
End Function

Private Function DateWithinReasonableRange(ByVal valueDate As Date) As Boolean
    DateWithinReasonableRange = (valueDate >= DateSerial(2000, 1, 1) And _
                                 valueDate <= DateSerial(2100, 12, 31))
End Function

Private Function CleanDateText(ByVal textValue As String) As String
    textValue = Replace(textValue, Chr$(160), " ")
    textValue = Replace(textValue, vbTab, " ")
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    CleanDateText = Kind_TrimAndCollapseSpaces(textValue)
End Function

' =============================================================================
' Generic helpers
' =============================================================================

Private Function DetermineSemesterFromKinder(ByVal wsKinder As Worksheet, _
                                             ByRef outMonth As Long, _
                                             ByRef outSemester As Long) As Boolean
    Dim referenceDate As Variant

    referenceDate = wsKinder.Range("T2").value
    If IsError(referenceDate) Then
        DetermineSemesterFromKinder = False
        Exit Function
    End If

    If Not IsDate(referenceDate) Then
        DetermineSemesterFromKinder = False
        Exit Function
    End If

    outMonth = Month(CDate(referenceDate))
    If outMonth >= 1 And outMonth <= 7 Then
        outSemester = 1
    Else
        outSemester = 2
    End If

    DetermineSemesterFromKinder = True
End Function

Private Function GetSubjectColumnForSemester(ByVal semester As Long) As String
    If semester = 1 Then
        GetSubjectColumnForSemester = KIND_KARTEI_COL_FACH_FIRST_HALF
    Else
        GetSubjectColumnForSemester = KIND_KARTEI_COL_FACH_SECOND_HALF
    End If
End Function

Private Function GetTeacherColumnForSemester(ByVal semester As Long) As String
    If semester = 1 Then
        GetTeacherColumnForSemester = COL_K
    Else
        GetTeacherColumnForSemester = COL_P
    End If
End Function

Private Function RowHasKinderIdentity(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    RowHasKinderIdentity = (Len(ValueTextOrEmpty(ws.Range(COL_B & rowNum).value)) > 0 And _
                            Len(ValueTextOrEmpty(ws.Range(COL_C & rowNum).value)) > 0 And _
                            Len(ValueTextOrEmpty(ws.Range(COL_E & rowNum).value)) > 0)
End Function

Private Function GetLastDataRow(ByVal ws As Worksheet) As Long
    Dim lastRowB As Long
    Dim lastRowC As Long

    lastRowB = ws.Cells(ws.rows.Count, COL_B).End(xlUp).row
    lastRowC = ws.Cells(ws.rows.Count, COL_C).End(xlUp).row

    If lastRowB > lastRowC Then
        GetLastDataRow = lastRowB
    Else
        GetLastDataRow = lastRowC
    End If
End Function

Private Function NormalizeKeyCell(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then
        NormalizeKeyCell = ""
    ElseIf IsNull(cellValue) Then
        NormalizeKeyCell = ""
    ElseIf IsEmpty(cellValue) Then
        NormalizeKeyCell = ""
    Else
        NormalizeKeyCell = LCase$(Kind_TrimAndCollapseSpaces(CStr(cellValue)))
    End If
End Function

Private Function IsActiveRecordSafe(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    Dim cellValue As Variant

    cellValue = ws.Range(COL_C & rowNum).value

    If IsError(cellValue) Then Exit Function
    If Len(ValueTextOrEmpty(cellValue)) = 0 Then Exit Function
    If ws.Range(COL_C & rowNum).Font.ColorIndex = 15 Then Exit Function

    IsActiveRecordSafe = True
End Function

Private Function SetCellIfDifferent(ByVal targetCell As Range, ByVal newValue As Variant) As Boolean
    If ValuesEqual(targetCell.value, newValue) Then
        SetCellIfDifferent = False
    Else
        targetCell.value = newValue
        SetCellIfDifferent = True
    End If
End Function

Private Function ValuesEqual(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
    If IsError(leftValue) Then
        ValuesEqual = False
        Exit Function
    End If

    If IsError(rightValue) Then
        ValuesEqual = False
        Exit Function
    End If

    If IsNull(leftValue) Then
        ValuesEqual = False
        Exit Function
    End If

    If IsNull(rightValue) Then
        ValuesEqual = False
        Exit Function
    End If

    If IsEmpty(leftValue) And IsEmpty(rightValue) Then
        ValuesEqual = True
        Exit Function
    End If

    If IsDate(leftValue) And IsDate(rightValue) Then
        ValuesEqual = (CLng(dateValue(CDate(leftValue))) = CLng(dateValue(CDate(rightValue))))
    Else
        ValuesEqual = (Trim$(CStr(leftValue)) = Trim$(CStr(rightValue)))
    End If
End Function

Private Function CellDateEquals(ByVal cellValue As Variant, ByVal compareDate As Date) As Boolean
    If IsError(cellValue) Then
        CellDateEquals = False
        Exit Function
    End If

    If Not IsDate(cellValue) Then
        CellDateEquals = False
    Else
        CellDateEquals = (CLng(dateValue(CDate(cellValue))) = CLng(dateValue(compareDate)))
    End If
End Function

Private Function IsBlankValue(ByVal cellValue As Variant) As Boolean
    If IsError(cellValue) Then
        IsBlankValue = True
    ElseIf IsNull(cellValue) Then
        IsBlankValue = True
    ElseIf IsEmpty(cellValue) Then
        IsBlankValue = True
    Else
        IsBlankValue = (Len(Trim$(CStr(cellValue))) = 0)
    End If
End Function

Private Function SafeValueText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then
        SafeValueText = "(Fehlerwert)"
    ElseIf IsNull(cellValue) Then
        SafeValueText = ""
    ElseIf IsEmpty(cellValue) Then
        SafeValueText = ""
    Else
        SafeValueText = Trim$(CStr(cellValue))
    End If
End Function

Private Function ValueTextOrEmpty(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then
        ValueTextOrEmpty = ""
    ElseIf IsNull(cellValue) Then
        ValueTextOrEmpty = ""
    ElseIf IsEmpty(cellValue) Then
        ValueTextOrEmpty = ""
    Else
        ValueTextOrEmpty = Trim$(CStr(cellValue))
    End If
End Function

Private Sub AddRowsToCollection(ByVal sourceRows As Collection, ByRef targetRows As Collection)
    Dim i As Long
    For i = 1 To sourceRows.Count
        targetRows.Add CLng(sourceRows(i))
    Next i
End Sub

Private Function DeleteRowsFromBottomUp(ByVal ws As Worksheet, ByVal rowsToDelete As Collection) As Long
    Dim sortedRows() As Long
    Dim i As Long
    Dim j As Long
    Dim tmp As Long
    Dim deletedCount As Long

    If rowsToDelete Is Nothing Then Exit Function
    If rowsToDelete.Count = 0 Then Exit Function

    ReDim sortedRows(1 To rowsToDelete.Count)
    For i = 1 To rowsToDelete.Count
        sortedRows(i) = CLng(rowsToDelete(i))
    Next i

    For i = 1 To UBound(sortedRows) - 1
        For j = i + 1 To UBound(sortedRows)
            If sortedRows(i) < sortedRows(j) Then
                tmp = sortedRows(i)
                sortedRows(i) = sortedRows(j)
                sortedRows(j) = tmp
            End If
        Next j
    Next i

    For i = 1 To UBound(sortedRows)
        ws.rows(sortedRows(i)).Delete
        deletedCount = deletedCount + 1
    Next i

    DeleteRowsFromBottomUp = deletedCount
End Function

Private Sub RenumberColumnA(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim rowNum As Long
    Dim counter As Long

    lastRow = GetLastDataRow(ws)
    If lastRow < BLANKS_FIRST_DATA_ROW Then Exit Sub

    counter = 1
    For rowNum = BLANKS_FIRST_DATA_ROW To lastRow
        If Len(ValueTextOrEmpty(ws.Range(COL_C & rowNum).value)) > 0 Then
            ws.Range(COL_A & rowNum).value = counter
            counter = counter + 1
        Else
            ws.Range(COL_A & rowNum).value = ""
        End If
    Next rowNum
End Sub

Private Sub SortByColumnC(ByVal ws As Worksheet)
    Dim lastRow As Long

    lastRow = GetLastDataRow(ws)
    If lastRow < BLANKS_FIRST_DATA_ROW Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range(COL_C & BLANKS_FIRST_DATA_ROW & ":" & COL_C & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SetRange ws.Range(COL_A & BLANKS_FIRST_DATA_ROW & ":" & COL_T & lastRow)
        .header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

' =============================================================================
' Logging and summary
' =============================================================================

Private Function EnsureErrorLogSheet() As Worksheet
    Dim ws As Worksheet

    On Error GoTo ErrHandler

    If Not Kind_TryGetWorksheet(ThisWorkbook, KIND_THIS_SHEET_ERRORLOG, ws) Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = KIND_THIS_SHEET_ERRORLOG
    End If

    Kind_EnsureErrorLogHeaders ws
    Set EnsureErrorLogSheet = ws
    Exit Function

ErrHandler:
    Set EnsureErrorLogSheet = Nothing
End Function

Private Sub LogIssue(ByRef stats As AdminSyncStats, _
                     ByVal familyID As String, _
                     ByVal lastName As String, _
                     ByVal firstName As String, _
                     ByVal subject As String, _
                     ByVal issueType As String, _
                     ByVal details As String)
    stats.WarningCount = stats.WarningCount + 1
    Kind_LogError familyID, lastName, firstName, subject, issueType, details
End Sub

Private Sub LogDuplicateKey(ByRef stats As AdminSyncStats, _
                            ByVal sheetName As String, _
                            ByVal key As String, _
                            ByVal rowsForKey As Collection)
    stats.DuplicateSkipped = stats.DuplicateSkipped + 1
    LogIssue stats, "", "", "", "", "DUPLICATE_KEY", _
             "Mehrere Zeilen auf '" & sheetName & "' fuer stabilen Schluessel (" & _
             key & "): " & CollectionRowsToText(rowsForKey)
End Sub

Private Function CollectionRowsToText(ByVal rowsForKey As Collection) As String
    Dim i As Long
    Dim result As String

    For i = 1 To rowsForKey.Count
        If Len(result) > 0 Then result = result & ", "
        result = result & CStr(rowsForKey(i))
    Next i

    CollectionRowsToText = result
End Function

Private Sub ShowSummary(ByRef stats As AdminSyncStats, _
                        ByVal monthNumber As Long, _
                        ByVal semester As Long)
    Dim msg As String

    msg = "Admin-Synchronisation abgeschlossen." & vbCrLf & vbCrLf
    msg = msg & "Monat aus Kinder!T2: " & monthNumber & vbCrLf
    msg = msg & "Semester: " & semester & vbCrLf & vbCrLf
    msg = msg & "Kinder geprueft: " & stats.KinderChecked & vbCrLf
    msg = msg & "Kinder geloescht (KN): " & stats.KinderDeletedKN & vbCrLf
    msg = msg & "Lehrer aktualisiert: " & stats.TeacherUpdated & vbCrLf
    msg = msg & "Datum H aktualisiert: " & stats.DateUpdated & vbCrLf & vbCrLf
    msg = msg & "Kinder_Blanks hinzugefuegt: " & stats.BlanksAdded & vbCrLf
    msg = msg & "Kinder_Blanks aktualisiert: " & stats.BlanksUpdated & vbCrLf
    msg = msg & "Kinder_Blanks geloescht: " & stats.BlanksDeleted & vbCrLf
    msg = msg & "Duplikate uebersprungen: " & stats.DuplicateSkipped & vbCrLf
    msg = msg & "Warnungen/ErrorLog: " & stats.WarningCount

    MsgBox msg, vbInformation, "Nachhilfe Synchronisation"
End Sub
