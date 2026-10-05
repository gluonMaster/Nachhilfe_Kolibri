Attribute VB_Name = "ValidateStudyHours_utils"
Option Explicit

Private Const VSH_LOG_PREFIX As String = "VSH_MissingPrev_"

Public Sub VSH_BuildPrevKeyIndex( _
    ByVal prevWs As Worksheet, _
    ByRef idKeys As Object, _
    ByRef childKeys As Object, _
    ByRef subjectKeys As Object, _
    ByRef fullKeys As Object)

    Set idKeys = CreateObject("Scripting.Dictionary")
    Set childKeys = CreateObject("Scripting.Dictionary")
    Set subjectKeys = CreateObject("Scripting.Dictionary")
    Set fullKeys = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = prevWs.Cells(prevWs.rows.Count, "A").End(xlUp).row
    If lastRow < 11 Then Exit Sub

    Dim rowNum As Long
    For rowNum = 11 To lastRow
        Dim idVal As String
        Dim surnameVal As String
        Dim nameVal As String
        Dim subjectVal As String
        Dim lessonTypeVal As String

        idVal = Trim(CStr(prevWs.Cells(rowNum, "A").value))
        If idVal = "" Then GoTo NextRow

        surnameVal = Trim(CStr(prevWs.Cells(rowNum, "B").value))
        nameVal = Trim(CStr(prevWs.Cells(rowNum, "C").value))
        subjectVal = Trim(CStr(prevWs.Cells(rowNum, "D").value))
        lessonTypeVal = Trim(CStr(prevWs.Cells(rowNum, "E").value))

        Dim idKey As String
        Dim childKey As String
        Dim subjectKey As String
        Dim fullKey As String

        idKey = idVal
        childKey = idVal & "|" & surnameVal & "|" & nameVal
        subjectKey = childKey & "|" & subjectVal
        fullKey = subjectKey & "|" & lessonTypeVal

        If Not idKeys.exists(idKey) Then idKeys.Add idKey, True
        If Not childKeys.exists(childKey) Then childKeys.Add childKey, True

        If subjectVal <> "" Then
            If Not subjectKeys.exists(subjectKey) Then subjectKeys.Add subjectKey, True
            If Not fullKeys.exists(fullKey) Then fullKeys.Add fullKey, True
        End If

NextRow:
    Next rowNum
End Sub

Public Function VSH_GetMissingMatchReason( _
    ByVal idKey As String, _
    ByVal childKey As String, _
    ByVal subjectKey As String, _
    ByVal fullKey As String, _
    ByVal idKeys As Object, _
    ByVal childKeys As Object, _
    ByVal subjectKeys As Object, _
    ByVal fullKeys As Object) As String

    If idKeys Is Nothing Or childKeys Is Nothing Or subjectKeys Is Nothing Or fullKeys Is Nothing Then
        VSH_GetMissingMatchReason = "Previous month lookup index is not initialized."
        Exit Function
    End If

    If subjectKeys.exists(subjectKey) Then
        VSH_GetMissingMatchReason = "ID, Surname, Name and Subject matched, but Lesson Type did not match."
    ElseIf childKeys.exists(childKey) Then
        VSH_GetMissingMatchReason = "ID, Surname and Name matched, but Subject did not match."
    ElseIf idKeys.exists(idKey) Then
        VSH_GetMissingMatchReason = "ID matched, but Surname/Name did not match."
    Else
        VSH_GetMissingMatchReason = "ID was not found on the previous month sheet."
    End If
End Function

Public Sub VSH_AddMissingPrevEntry( _
    ByRef entries As Collection, _
    ByVal sourceRow As Long, _
    ByVal childID As String, _
    ByVal surnameVal As String, _
    ByVal nameVal As String, _
    ByVal subjectVal As String, _
    ByVal lessonTypeVal As String, _
    ByVal reasonText As String)

    If entries Is Nothing Then Set entries = New Collection

    Dim rowItem(1 To 7) As Variant
    rowItem(1) = sourceRow
    rowItem(2) = childID
    rowItem(3) = surnameVal
    rowItem(4) = nameVal
    rowItem(5) = subjectVal
    rowItem(6) = lessonTypeVal
    rowItem(7) = reasonText

    entries.Add rowItem
End Sub

Public Function VSH_CreateMissingPrevLogSheet(ByVal currentWs As Worksheet) As Worksheet
    Dim wb As Workbook
    Set wb = currentWs.Parent

    Dim logSheetName As String
    logSheetName = VSH_BuildLogSheetName(currentWs.name)

    On Error GoTo ErrHandler

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(logSheetName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Dim logWs As Worksheet
    Set logWs = wb.Worksheets.Add(After:=currentWs)
    logWs.name = logSheetName

    logWs.Range("A1").value = "ValidateStudyHours: Potential Missing Matches from Previous Month"
    logWs.Range("A2").value = "Current month sheet:"
    logWs.Range("A3").value = "Previous month sheet:"
    logWs.Range("A4").value = "Generated at:"
    logWs.Range("A5").value = "This sheet is temporary. Use the button to delete it."

    logWs.Range("A7").value = "Source Row"
    logWs.Range("B7").value = "ID"
    logWs.Range("C7").value = "Surname"
    logWs.Range("D7").value = "Name"
    logWs.Range("E7").value = "Subject"
    logWs.Range("F7").value = "Lesson Type"
    logWs.Range("G7").value = "What was not found"

    With logWs.Range("A1:G1")
        .Font.Bold = True
        .Font.Size = 12
    End With

    With logWs.Range("A7:G7")
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
    End With

    VSH_AddDeleteLogButton logWs

    Set VSH_CreateMissingPrevLogSheet = logWs
    Exit Function

ErrHandler:
    Application.DisplayAlerts = True
    MsgBox "Unable to create temporary missing-match log sheet: " & Err.Description, vbCritical
End Function

Public Sub VSH_WriteMissingPrevEntries( _
    ByVal logWs As Worksheet, _
    ByVal currentWs As Worksheet, _
    ByVal prevWs As Worksheet, _
    ByVal entries As Collection)

    If logWs Is Nothing Then Exit Sub
    If entries Is Nothing Then Exit Sub
    If entries.Count = 0 Then Exit Sub

    logWs.Range("B2").value = currentWs.name
    logWs.Range("B3").value = prevWs.name
    logWs.Range("B4").value = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    Dim dataArr() As Variant
    ReDim dataArr(1 To entries.Count, 1 To 7)

    Dim i As Long
    For i = 1 To entries.Count
        Dim itemArr As Variant
        itemArr = entries(i)

        dataArr(i, 1) = itemArr(1)
        dataArr(i, 2) = itemArr(2)
        dataArr(i, 3) = itemArr(3)
        dataArr(i, 4) = itemArr(4)
        dataArr(i, 5) = itemArr(5)
        dataArr(i, 6) = itemArr(6)
        dataArr(i, 7) = itemArr(7)
    Next i

    logWs.Range("A8").Resize(entries.Count, 7).value = dataArr

    logWs.Columns("A:G").AutoFit
    logWs.Range("A7:G7").AutoFilter
End Sub

Public Sub VSH_DeleteActiveLogSheet()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    If Left$(ws.name, Len(VSH_LOG_PREFIX)) <> VSH_LOG_PREFIX Then
        MsgBox "This button can only delete ValidateStudyHours temporary log sheets.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Delete this temporary log sheet?", vbQuestion + vbYesNo, "Delete Log Sheet") <> vbYes Then
        Exit Sub
    End If

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    MsgBox "Unable to delete the log sheet: " & Err.Description, vbCritical
End Sub

Private Sub VSH_AddDeleteLogButton(ByVal logWs As Worksheet)
    On Error GoTo ErrHandler

    On Error Resume Next
    logWs.Buttons("VSH_DeleteLogButton").Delete
    On Error GoTo ErrHandler

    Dim deleteBtn As Button
    Set deleteBtn = logWs.Buttons.Add(logWs.Range("I1").Left, logWs.Range("I1").Top, 230, 24)
    Dim workbookNameEscaped As String
    workbookNameEscaped = Replace(ThisWorkbook.name, "'", "''")

    With deleteBtn
        .name = "VSH_DeleteLogButton"
        .Caption = "Delete this temporary log sheet"
        .OnAction = "'" & workbookNameEscaped & "'!VSH_DeleteActiveLogSheet"
    End With
    Exit Sub

ErrHandler:
    MsgBox "Unable to create delete button on temporary log sheet: " & Err.Description, vbExclamation
End Sub

Private Function VSH_BuildLogSheetName(ByVal monthSheetName As String) As String
    Dim rawName As String
    rawName = VSH_LOG_PREFIX & monthSheetName

    If Len(rawName) <= 31 Then
        VSH_BuildLogSheetName = rawName
    Else
        VSH_BuildLogSheetName = Left$(rawName, 31)
    End If
End Function
