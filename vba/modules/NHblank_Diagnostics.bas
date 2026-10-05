Attribute VB_Name = "NHblank_Diagnostics"
Option Explicit

' Comprehensive diagnostic check for the entire system
Public Sub RunFullDiagnostics()
    Dim msg As String
    Dim issues As Long
    
    issues = 0
    msg = "=== SYSTEM DIAGNOSTICS ===" & vbCrLf & vbCrLf
    
    ' 1. Check if Kinder worksheet exists (source for Kinder_Blanks)
    msg = msg & "1. Worksheet 'Kinder' (source):" & vbCrLf
    If CheckKinderWorksheet(msg) Then
        msg = msg & "   OK - Worksheet exists" & vbCrLf
    Else
        msg = msg & "   ERROR - Worksheet not found!" & vbCrLf
        issues = issues + 1
    End If
    msg = msg & vbCrLf
    
    ' 1b. Check if Kinder_Blanks worksheet exists or can be created
    msg = msg & "1b. Worksheet 'Kinder_Blanks':" & vbCrLf
    If CheckKinderBlanksWorksheet(msg) Then
        msg = msg & "   OK - Worksheet exists or can be created" & vbCrLf
    Else
        msg = msg & "   ERROR - Worksheet not available and cannot be created!" & vbCrLf
        issues = issues + 1
    End If
    msg = msg & vbCrLf
    
    ' 2. Check current position
    msg = msg & "2. Current Position:" & vbCrLf
    If ActiveSheet.name = "Kinder_Blanks" Or ActiveSheet.name = "Kinder" Then
        msg = msg & "   OK - On '" & ActiveSheet.name & "' sheet" & vbCrLf
        msg = msg & "   Current row: " & ActiveCell.row & vbCrLf
        msg = msg & "   Current column: " & ActiveCell.Column & vbCrLf
    Else
        msg = msg & "   WARNING - Not on 'Kinder_Blanks' or 'Kinder' sheet!" & vbCrLf
        msg = msg & "   Current sheet: " & ActiveSheet.name & vbCrLf
        issues = issues + 1
    End If
    msg = msg & vbCrLf
    
    ' 3. Check active records
    msg = msg & "3. Active Records Analysis:" & vbCrLf
    CheckActiveRecords msg, issues
    msg = msg & vbCrLf
    
    ' 4. Check current row data
    If (ActiveSheet.name = "Kinder_Blanks" Or ActiveSheet.name = "Kinder") And ActiveCell.row >= 5 Then
        msg = msg & "4. Current Row Data:" & vbCrLf
        CheckCurrentRowData msg, issues
        msg = msg & vbCrLf
    End If
    
    ' 5. BG normalization self-test
    msg = msg & "5. BG Number Normalization:" & vbCrLf
    Dim bgSelfTest As String
    bgSelfTest = NHblank_BgNummer.NHblank_BgNormalizationSelfTest()
    If bgSelfTest = "OK" Then
        msg = msg & "   OK - " & bgSelfTest & vbCrLf
    Else
        msg = msg & "   ERROR - " & bgSelfTest & vbCrLf
        issues = issues + 1
    End If
    msg = msg & vbCrLf

    ' 6. Summary
    msg = msg & "=== SUMMARY ===" & vbCrLf
    If issues = 0 Then
        msg = msg & "No critical issues found." & vbCrLf
    Else
        msg = msg & "Found " & issues & " issue(s) - see details above." & vbCrLf
    End If
    
    MsgBox msg, vbInformation, "System Diagnostics Report"
End Sub

' Check if Kinder worksheet exists
Private Function CheckKinderWorksheet(ByRef msg As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Kinder")
    On Error GoTo 0
    
    CheckKinderWorksheet = Not (ws Is Nothing)
End Function

' Check if Kinder_Blanks worksheet exists or can be created
Private Function CheckKinderBlanksWorksheet(ByRef msg As String) As Boolean
    Dim ws As Worksheet
    
    ' First check if Kinder_Blanks already exists
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Kinder_Blanks")
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        CheckKinderBlanksWorksheet = True
        Exit Function
    End If
    
    ' If not exists, check if Kinder exists (can create from it)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Kinder")
    On Error GoTo 0
    
    CheckKinderBlanksWorksheet = Not (ws Is Nothing)
End Function

' Check and count active records
Private Sub CheckActiveRecords(ByRef msg As String, ByRef issues As Long)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim activeCount As Long
    Dim inactiveCount As Long
    Dim emptyCount As Long
    
    ' Try Kinder_Blanks first, fall back to Kinder
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Kinder_Blanks")
    On Error GoTo 0
    
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets("Kinder")
        On Error GoTo 0
    End If
    
    If ws Is Nothing Then
        msg = msg & "   Cannot check - worksheet not found" & vbCrLf
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.rows.Count, "C").End(xlUp).row
    msg = msg & "   Last row with data: " & lastRow & vbCrLf
    
    activeCount = 0
    inactiveCount = 0
    emptyCount = 0
    
    For i = 5 To lastRow
        If ws.Cells(i, "C").value = "" Then
            emptyCount = emptyCount + 1
        ElseIf ws.Cells(i, "C").Font.ColorIndex = 15 Then
            inactiveCount = inactiveCount + 1
        Else
            activeCount = activeCount + 1
        End If
    Next i
    
    msg = msg & "   Active records (ColorIndex <> 15): " & activeCount & vbCrLf
    msg = msg & "   Inactive records (ColorIndex = 15): " & inactiveCount & vbCrLf
    msg = msg & "   Empty records: " & emptyCount & vbCrLf
    
    If activeCount = 0 Then
        msg = msg & "   WARNING - No active records found!" & vbCrLf
        issues = issues + 1
    End If
End Sub

' Check current row data completeness
Private Sub CheckCurrentRowData(ByRef msg As String, ByRef issues As Long)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastName As String
    Dim firstName As String
    Dim discipline As String
    Dim dateFrom As Variant
    Dim dateTo As Variant
    Dim referenceDate As Variant
    Dim birthDateValue As Variant
    Dim normalizedBg As String
    Dim isJobcenter As Boolean
    Dim bgError As String
    Dim hasIssues As Boolean
    
    Set ws = ActiveSheet
    rowNum = ActiveCell.row
    hasIssues = False
    
    lastName = Trim(ws.Cells(rowNum, "C").value)
    firstName = Trim(ws.Cells(rowNum, "D").value)
    discipline = Trim(ws.Cells(rowNum, "E").value)
    dateFrom = ws.Cells(rowNum, "G").value
    dateTo = ws.Cells(rowNum, "H").value
    referenceDate = ws.Range("T2").value
    birthDateValue = ws.Cells(rowNum, "L").value
    
    msg = msg & "   Row: " & rowNum & vbCrLf
    msg = msg & "   ColorIndex: " & ws.Cells(rowNum, "C").Font.ColorIndex & vbCrLf
    
    If lastName = "" Then
        msg = msg & "   ERROR - Last name (column C) is empty!" & vbCrLf
        hasIssues = True
    Else
        msg = msg & "   Last name: " & lastName & vbCrLf
    End If
    
    If firstName = "" Then
        msg = msg & "   ERROR - First name (column D) is empty!" & vbCrLf
        hasIssues = True
    Else
        msg = msg & "   First name: " & firstName & vbCrLf
    End If
    
    If discipline = "" Then
        msg = msg & "   ERROR - Discipline (column E) is empty!" & vbCrLf
        hasIssues = True
    Else
        msg = msg & "   Discipline: " & discipline & vbCrLf
    End If
    
    ' Check dates
    If IsEmpty(dateFrom) Then
        msg = msg & "   Date From (G): Empty" & vbCrLf
    ElseIf Not IsDate(dateFrom) Then
        msg = msg & "   ERROR - Date From (G) is not a valid date: " & dateFrom & vbCrLf
        hasIssues = True
    Else
        msg = msg & "   Date From (G): " & Format(dateFrom, "dd.mm.yyyy") & vbCrLf
    End If
    
    If IsEmpty(dateTo) Then
        msg = msg & "   Date To (H): Empty" & vbCrLf
    ElseIf Not IsDate(dateTo) Then
        msg = msg & "   ERROR - Date To (H) is not a valid date: " & dateTo & vbCrLf
        hasIssues = True
    Else
        msg = msg & "   Date To (H): " & Format(dateTo, "dd.mm.yyyy") & vbCrLf
    End If
    
    If Not IsDate(referenceDate) Then
        msg = msg & "   ERROR - Reference Date (T2) is not a valid date: " & referenceDate & vbCrLf
        hasIssues = True
    Else
        msg = msg & "   Reference Date (T2): " & Format(referenceDate, "dd.mm.yyyy") & vbCrLf
    End If

    If Not IsDate(birthDateValue) Then
        msg = msg & "   ERROR - Birth date (L) is not valid!" & vbCrLf
        hasIssues = True
    Else
        msg = msg & "   Birth date (L): valid" & vbCrLf
    End If

    If Not NHblank_BgNummer.NHblank_TryNormalizeBgCell( _
        ws.Cells(rowNum, "O"), normalizedBg, isJobcenter, bgError) Then
        msg = msg & "   ERROR - BG number (O): " & bgError & vbCrLf
        hasIssues = True
    ElseIf isJobcenter Then
        msg = msg & "   BG number (O): valid Jobcenter format" & vbCrLf
    Else
        msg = msg & "   BG number (O): valid Sozialamt format" & vbCrLf
    End If
    
    If hasIssues Then
        issues = issues + 1
    End If
End Sub

' Check if the legacy Sozialamt Excel template is valid.
Public Function CheckLegacyTemplateFile(ByVal templatePath As String, _
                                        ByRef errorMsg As String) As Boolean
    Dim wbTemplate As Workbook
    Dim wsTemplate As Worksheet
    
    On Error GoTo ErrorHandler
    
    ' Check if file exists
    If Dir(templatePath) = "" Then
        errorMsg = "Template file not found: " & templatePath
        CheckLegacyTemplateFile = False
        Exit Function
    End If
    
    ' Try to open template
    Set wbTemplate = Workbooks.Open(templatePath, ReadOnly:=True)
    
    ' Check if Muster worksheet exists
    On Error Resume Next
    Set wsTemplate = wbTemplate.Worksheets("Muster")
    On Error GoTo ErrorHandler
    
    If wsTemplate Is Nothing Then
        errorMsg = "Worksheet 'Muster' not found in template file!"
        wbTemplate.Close SaveChanges:=False
        CheckLegacyTemplateFile = False
        Exit Function
    End If
    
    ' All checks passed
    wbTemplate.Close SaveChanges:=False
    CheckLegacyTemplateFile = True
    Exit Function
    
ErrorHandler:
    errorMsg = "Error checking template: " & Err.Description
    If Not wbTemplate Is Nothing Then
        wbTemplate.Close SaveChanges:=False
    End If
    CheckLegacyTemplateFile = False
End Function

' Backward-compatible alias.
Public Function CheckTemplateFile(ByVal templatePath As String, _
                                  ByRef errorMsg As String) As Boolean
    CheckTemplateFile = CheckLegacyTemplateFile(templatePath, errorMsg)
End Function

' Check the new Word template, all required content-control tags and page count.
Public Function CheckWordTemplateFile(ByVal templatePath As String, _
                                      ByRef errorMsg As String) As Boolean
    CheckWordTemplateFile = _
        NHblank_WordTemplate.NHblank_CheckWordTemplate(templatePath, errorMsg)
End Function

' Check if target folder is writable
Public Function CheckTargetFolder(ByVal folderPath As String, ByRef errorMsg As String) As Boolean
    Dim testFile As String
    Dim fso As Object
    
    On Error GoTo ErrorHandler
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if folder exists
    If Not fso.FolderExists(folderPath) Then
        errorMsg = "Target folder does not exist: " & folderPath
        CheckTargetFolder = False
        Exit Function
    End If
    
    ' Try to create a test file to check write permissions
    testFile = folderPath & "\_test_write_" & Format(Now, "yyyymmddhhnnss") & ".tmp"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open testFile For Output As #fileNum
    Print #fileNum, "test"
    Close #fileNum
    
    ' Delete test file
    Kill testFile
    
    CheckTargetFolder = True
    Exit Function
    
ErrorHandler:
    errorMsg = "Cannot write to target folder: " & Err.Description
    CheckTargetFolder = False
End Function

' Detailed check before processing
Public Function PreProcessCheck( _
    ByVal wordTemplatePath As String, _
    ByVal legacyTemplatePath As String, _
    ByVal targetFolder As String, _
    ByRef errorMsg As String _
) As Boolean
    Dim ws As Worksheet
    Dim wsKinder As Worksheet
    
    ' Check if Kinder_Blanks exists or can be created from Kinder
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Kinder_Blanks")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Check if Kinder exists (required to create Kinder_Blanks)
        On Error Resume Next
        Set wsKinder = ThisWorkbook.Worksheets("Kinder")
        On Error GoTo 0
        
        If wsKinder Is Nothing Then
            errorMsg = "Worksheet 'Kinder_Blanks' not found and 'Kinder' is missing - cannot create!"
            PreProcessCheck = False
            Exit Function
        End If
        ' Kinder exists, so Kinder_Blanks can be auto-created during processing
    End If
    
    ' Check the new Word template.
    If Not CheckWordTemplateFile(wordTemplatePath, errorMsg) Then
        PreProcessCheck = False
        Exit Function
    End If

    ' Check the legacy Excel template required for Sozialamt records.
    If Not CheckLegacyTemplateFile(legacyTemplatePath, errorMsg) Then
        PreProcessCheck = False
        Exit Function
    End If
    
    ' Check target folder
    If Not CheckTargetFolder(targetFolder, errorMsg) Then
        PreProcessCheck = False
        Exit Function
    End If
    
    PreProcessCheck = True
End Function
