Attribute VB_Name = "ModuleManager"
'==========================
'   Module: ModuleManager
'   Purpose: Import/Export VBA modules (.bas, .cls) with user-friendly dialogs
'
'   Features:
'   - Import multiple .bas/.cls/.frm files at once (a form needs its .frx beside it)
'   - Warning if modules with same names already exist
'   - Handles files without proper Attribute headers or with LF line endings
'   - Replaces a module only after the new one was imported
'   - Refuses code of sheets and the workbook (paste it in the VBA editor)
'   - Export all modules to selected folder
'   - Export selected modules from a list
'
'   Entry Points:
'   - ImportModules()        - Select and import multiple .bas/.cls/.frm files
'   - ExportAllModules()     - Export all modules to selected folder
'   - ExportSelectedModules() - Show list and export selected modules
'
'   Requirements:
'   - Trust access to VBA project object model must be enabled
'     (File > Options > Trust Center > Trust Center Settings > Macro Settings)
'==========================

Option Explicit

' Module type constants
Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100

' ============================================================
' IMPORT MODULES
' ============================================================

' Main entry point: Import multiple .bas/.cls/.frm files
Public Sub ImportModules()
    On Error GoTo ErrorHandler
    
    ' Check VBA project access
    If Not HasVBProjectAccess() Then
        MsgBox "Zugriff auf VBA-Projekt nicht moeglich." & vbCrLf & vbCrLf & _
               "Bitte aktivieren Sie:" & vbCrLf & _
               "Datei > Optionen > Trust Center > Einstellungen fuer das Trust Center > Makroeinstellungen" & vbCrLf & _
               "> 'Zugriff auf das VBA-Projektobjektmodell vertrauen'", _
               vbCritical, "VBA-Projektzugriff erforderlich"
        Exit Sub
    End If
    
    ' Open file dialog for multiple selection
    Dim fd As fileDialog
    Set fd = Application.fileDialog(msoFileDialogFilePicker)
    
    With fd
        .title = "VBA-Module zum Importieren auswaehlen"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "VBA-Module", "*.bas;*.cls;*.frm"
        .Filters.Add "Standard-Module (*.bas)", "*.bas"
        .Filters.Add "Klassen-Module (*.cls)", "*.cls"
        .Filters.Add "Formulare (*.frm, mit .frx daneben)", "*.frm"
        .Filters.Add "Alle Dateien", "*.*"
        .FilterIndex = 1
        
        If .Show <> -1 Then
            ' User cancelled
            Exit Sub
        End If
        
        If .SelectedItems.Count = 0 Then
            Exit Sub
        End If
        
        ' Collect files to import
        Dim filePaths() As String
        ReDim filePaths(1 To .SelectedItems.Count)
        
        Dim i As Long
        For i = 1 To .SelectedItems.Count
            filePaths(i) = .SelectedItems(i)
        Next i
    End With
    
    ' Check for existing modules and confirm
    Dim existingModules As String
    existingModules = CheckExistingModules(filePaths)
    
    If existingModules <> "" Then
        Dim result As VbMsgBoxResult
        result = MsgBox("Folgende Module existieren bereits und werden ueberschrieben:" & vbCrLf & vbCrLf & _
                       existingModules & vbCrLf & _
                       "Moechten Sie fortfahren?", _
                       vbYesNo + vbExclamation, "Module existieren bereits")
        
        If result = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Import files
    Dim importedCount As Long
    Dim failedCount As Long
    Dim failedFiles As String
    
    Application.ScreenUpdating = False
    
    Dim reason As String
    For i = LBound(filePaths) To UBound(filePaths)
        reason = ImportSingleModule(filePaths(i))
        If reason = "" Then
            importedCount = importedCount + 1
        Else
            failedCount = failedCount + 1
            failedFiles = failedFiles & vbCrLf & "  - " & GetFileName(filePaths(i)) & ": " & reason
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    ' Show result
    Dim msg As String
    msg = "Import abgeschlossen:" & vbCrLf & vbCrLf & _
          "Erfolgreich importiert: " & importedCount & " Modul(e)"
    
    If failedCount > 0 Then
        msg = msg & vbCrLf & "Fehlgeschlagen: " & failedCount & " Modul(e)" & failedFiles
    End If
    
    MsgBox msg, IIf(failedCount > 0, vbExclamation, vbInformation), "Import-Ergebnis"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Importieren: " & Err.Description, vbCritical, "Import-Fehler"
End Sub

' Check which modules already exist
Private Function CheckExistingModules(ByRef filePaths() As String) As String
    Dim result As String
    result = ""
    
    Dim i As Long
    For i = LBound(filePaths) To UBound(filePaths)
        Dim moduleName As String
        moduleName = GetModuleNameFromFile(filePaths(i))
        
        If ModuleExists(moduleName) Then
            result = result & vbCrLf & "  - " & moduleName
        End If
    Next i
    
    CheckExistingModules = result
End Function

' Get module name from file (reads Attribute VB_Name or uses filename)
Private Function GetModuleNameFromFile(ByVal filePath As String) As String
    On Error GoTo UseFileName
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1, False) ' 1 = ForReading
    
    Dim line As String
    Dim moduleName As String
    moduleName = ""
    
    ' Read first 20 lines looking for Attribute VB_Name
    Dim lineCount As Long
    lineCount = 0
    
    Do While Not ts.AtEndOfStream And lineCount < 20
        line = ts.ReadLine
        lineCount = lineCount + 1
        
        If InStr(1, line, "Attribute VB_Name", vbTextCompare) > 0 Then
            ' Extract name from: Attribute VB_Name = "ModuleName"
            Dim startPos As Long, endPos As Long
            startPos = InStr(line, """")
            If startPos > 0 Then
                endPos = InStr(startPos + 1, line, """")
                If endPos > startPos Then
                    moduleName = mid(line, startPos + 1, endPos - startPos - 1)
                End If
            End If
            Exit Do
        End If
    Loop
    
    ts.Close
    
    If moduleName = "" Then
        GoTo UseFileName
    End If
    
    GetModuleNameFromFile = moduleName
    Exit Function
    
UseFileName:
    ' Use filename without extension as module name
    GetModuleNameFromFile = GetFileNameWithoutExtension(filePath)
End Function

' Check if module with given name exists
Private Function ModuleExists(ByVal moduleName As String) As Boolean
    On Error Resume Next
    
    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
    
    ModuleExists = Not (vbComp Is Nothing)
    
    On Error GoTo 0
End Function

' Import a single module file. Returns "" on success, otherwise the reason.
Private Function ImportSingleModule(ByVal filePath As String) As String
    Dim moduleName As String, tempPath As String, oldComp As Object
    On Error GoTo Failed
    
    moduleName = GetModuleNameFromFile(filePath)
    If StrComp(moduleName, "ModuleManager", vbTextCompare) = 0 Then
        ImportSingleModule = "ModuleManager kann sich nicht selbst ersetzen; bitte im VBA-Editor manuell importieren"
        Exit Function
    End If
    
    ' Sheet and workbook modules cannot be removed: the old one would stay renamed next to
    ' a class module that never runs.
    If ModuleExists(moduleName) Then
        If ThisWorkbook.VBProject.VBComponents(moduleName).Type = vbext_ct_Document Then
            ImportSingleModule = "Code von Tabellen und Arbeitsmappe bitte im VBA-Editor einfuegen"
            Exit Function
        End If
    End If
    
    ' Import never reads the selected file directly: a class module is recognised only with
    ' CRLF line endings, otherwise it arrives as a standard module with its header as code.
    Dim reason As String
    tempPath = CreateImportCopy(filePath, moduleName, reason)
    If tempPath = "" Then
        ImportSingleModule = reason
        Exit Function
    End If
    
    ' A removed module keeps its name until the macro ends, so the new one would get a
    ' suffixed name. The old module steps aside and is removed only after a successful import.
    If ModuleExists(moduleName) Then
        Set oldComp = ThisWorkbook.VBProject.VBComponents(moduleName)
        oldComp.name = FreeModuleName(moduleName)
    End If
    ThisWorkbook.VBProject.VBComponents.Import tempPath
    If Not oldComp Is Nothing Then ThisWorkbook.VBProject.VBComponents.Remove oldComp
    ImportSingleModule = ""
    
Done:
    On Error Resume Next
    If Len(tempPath) > 0 Then DeleteImportFolder tempPath
    Exit Function
    
Failed:
    ImportSingleModule = "Fehler " & Err.Number & ": " & Err.Description
    Debug.Print "Error importing " & filePath & ": " & Err.Description
    Resume KeepOld
KeepOld:
    ' Without a new module the old one keeps its name and code.
    On Error Resume Next
    If Not oldComp Is Nothing Then
        If Not ModuleExists(moduleName) Then oldComp.name = moduleName
    End If
    GoTo Done
End Function

' A name for a module that is about to be replaced. A removed form keeps its name until the
' workbook is saved, so the name is new for the whole session, not only for the project.
Private Function FreeModuleName(ByVal baseName As String) As String
    Static serial As Long
    Do
        serial = serial + 1
        FreeModuleName = Left$(baseName, 18) & "_old" & Format$(Now, "hhnnss") & serial
    Loop While ModuleExists(FreeModuleName)
End Function

' Temp copy for VBComponents.Import in a folder of its own: CRLF line endings, an Attribute
' header, and for a form its binary .frx under the name the .frm refers to.
' Returns the path of the copy, or "" with the reason.
Private Function CreateImportCopy(ByVal sourcePath As String, ByVal moduleName As String, _
                                  ByRef reason As String) As String
    Dim fso As Object, ts As Object, ext As String, content As String, header As String
    Dim folder As String, tempPath As String, frxName As String, frxSource As String
    Dim pos As Long, quotePos As Long
    On Error GoTo ErrorHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' A file of another type is imported as a standard module.
    ext = LCase$(fso.GetExtensionName(sourcePath))
    If ext <> "cls" And ext <> "frm" Then ext = "bas"
    
    Set ts = fso.OpenTextFile(sourcePath, 1, False)
    content = ts.ReadAll
    ts.Close
    ' Any mix of CR, LF and CRLF becomes CRLF.
    content = Replace(Replace(content, vbCrLf, vbLf), vbCr, vbLf)
    content = Replace(content, vbLf, vbCrLf)
    
    folder = fso.BuildPath(Environ("TEMP"), "VbaImport_" & fso.GetBaseName(fso.GetTempName()))
    fso.CreateFolder folder
    tempPath = fso.BuildPath(folder, moduleName & "." & ext)
    
    If ext = "frm" Then
        ' The designer of a form is the binary .frx, which Import looks up beside the .frm.
        If InStr(1, Left$(content, 2000), "Attribute VB_Name", vbTextCompare) = 0 Then
            reason = "Formulardatei ohne Kopfzeilen"
            GoTo Refused
        End If
        pos = InStr(1, content, ".frx""", vbTextCompare)
        If pos > 0 Then
            quotePos = InStrRev(content, """", pos)
            frxName = mid$(content, quotePos + 1, pos + 3 - quotePos)
            frxSource = fso.BuildPath(fso.GetParentFolderName(sourcePath), frxName)
            If Not fso.fileExists(frxSource) Then _
                frxSource = fso.BuildPath(fso.GetParentFolderName(sourcePath), fso.GetBaseName(sourcePath) & ".frx")
            If Not fso.fileExists(frxSource) Then
                reason = "Die zugehoerige Datei " & frxName & " fehlt neben der .frm-Datei"
                GoTo Refused
            End If
            fso.CopyFile frxSource, fso.BuildPath(folder, frxName)
        End If
    Else
        If InStr(1, Left$(content, 2000), "Attribute VB_Name", vbTextCompare) = 0 Then
            If ext = "cls" Then
                header = "VERSION 1.0 CLASS" & vbCrLf & _
                         "BEGIN" & vbCrLf & _
                         "  MultiUse = -1  'True" & vbCrLf & _
                         "END" & vbCrLf & _
                         "Attribute VB_Name = """ & moduleName & """" & vbCrLf & _
                         "Attribute VB_GlobalNameSpace = False" & vbCrLf & _
                         "Attribute VB_Creatable = False" & vbCrLf & _
                         "Attribute VB_PredeclaredId = False" & vbCrLf & _
                         "Attribute VB_Exposed = False" & vbCrLf
            Else
                header = "Attribute VB_Name = """ & moduleName & """" & vbCrLf
            End If
        End If
    End If
    
    Set ts = fso.CreateTextFile(tempPath, True)
    ts.Write header & content
    ts.Close
    CreateImportCopy = tempPath
    Exit Function
    
ErrorHandler:
    reason = "Datei konnte nicht vorbereitet werden (" & Err.Description & ")"
    Debug.Print "Error creating import copy: " & Err.Description
    Resume Refused
Refused:
    If Len(folder) > 0 Then DeleteImportFolder fso.BuildPath(folder, "x")
    CreateImportCopy = ""
End Function

' Removes the temp folder that holds an import copy
Private Sub DeleteImportFolder(ByVal tempPath As String)
    Dim fso As Object, folder As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.GetParentFolderName(tempPath)
    If InStr(1, fso.GetFileName(folder), "VbaImport_", vbTextCompare) = 1 Then fso.DeleteFolder folder, True
End Sub

' ============================================================
' EXPORT ALL MODULES
' ============================================================

' Export all .bas and .cls modules to selected folder
Public Sub ExportAllModules()
    On Error GoTo ErrorHandler
    
    ' Check VBA project access
    If Not HasVBProjectAccess() Then
        MsgBox "Zugriff auf VBA-Projekt nicht moeglich." & vbCrLf & vbCrLf & _
               "Bitte aktivieren Sie 'Zugriff auf das VBA-Projektobjektmodell vertrauen'", _
               vbCritical, "VBA-Projektzugriff erforderlich"
        Exit Sub
    End If
    
    ' Select target folder
    Dim targetFolder As String
    targetFolder = BrowseForFolder("Zielordner fuer Export auswaehlen")
    
    If targetFolder = "" Then
        Exit Sub
    End If
    
    ' Ensure trailing backslash
    If Right(targetFolder, 1) <> "\" Then
        targetFolder = targetFolder & "\"
    End If
    
    ' Export all eligible modules
    Dim exportedCount As Long
    exportedCount = 0
    
    Application.ScreenUpdating = False
    
    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If IsExportableModule(vbComp) Then
            Dim filePath As String
            filePath = targetFolder & vbComp.name & GetModuleExtension(vbComp)
            
            vbComp.Export filePath
            exportedCount = exportedCount + 1
        End If
    Next vbComp
    
    Application.ScreenUpdating = True
    
    MsgBox "Export abgeschlossen:" & vbCrLf & vbCrLf & _
           "Exportiert: " & exportedCount & " Modul(e)" & vbCrLf & _
           "Zielordner: " & targetFolder, _
           vbInformation, "Export-Ergebnis"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Exportieren: " & Err.Description, vbCritical, "Export-Fehler"
End Sub

' ============================================================
' EXPORT SELECTED MODULES
' ============================================================

' Show list of modules and export selected ones
Public Sub ExportSelectedModules()
    On Error GoTo ErrorHandler
    
    ' Check VBA project access
    If Not HasVBProjectAccess() Then
        MsgBox "Zugriff auf VBA-Projekt nicht moeglich." & vbCrLf & vbCrLf & _
               "Bitte aktivieren Sie 'Zugriff auf das VBA-Projektobjektmodell vertrauen'", _
               vbCritical, "VBA-Projektzugriff erforderlich"
        Exit Sub
    End If
    
    ' Get list of exportable modules
    Dim moduleList As String
    Dim moduleCount As Long
    moduleCount = 0
    
    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If IsExportableModule(vbComp) Then
            moduleCount = moduleCount + 1
            moduleList = moduleList & moduleCount & ". " & vbComp.name & _
                        " (" & GetModuleTypeName(vbComp) & ")" & vbCrLf
        End If
    Next vbComp
    
    If moduleCount = 0 Then
        MsgBox "Keine exportierbaren Module gefunden.", vbInformation, "Keine Module"
        Exit Sub
    End If
    
    ' Ask user to select modules (by number, comma-separated)
    Dim selection As String
    selection = InputBox("Verfuegbare Module:" & vbCrLf & vbCrLf & _
                        moduleList & vbCrLf & _
                        "Geben Sie die Nummern der zu exportierenden Module ein" & vbCrLf & _
                        "(kommagetrennt, z.B. '1,3,5' oder 'alle' fuer alle):", _
                        "Module zum Exportieren auswaehlen", "alle")
    
    If selection = "" Then
        Exit Sub
    End If
    
    ' Select target folder
    Dim targetFolder As String
    targetFolder = BrowseForFolder("Zielordner fuer Export auswaehlen")
    
    If targetFolder = "" Then
        Exit Sub
    End If
    
    ' Ensure trailing backslash
    If Right(targetFolder, 1) <> "\" Then
        targetFolder = targetFolder & "\"
    End If
    
    ' Parse selection
    Dim selectedNumbers() As String
    Dim exportAll As Boolean
    exportAll = (LCase(Trim(selection)) = "alle" Or LCase(Trim(selection)) = "all")
    
    If Not exportAll Then
        selectedNumbers = Split(selection, ",")
    End If
    
    ' Export selected modules
    Dim exportedCount As Long
    exportedCount = 0
    Dim currentIndex As Long
    currentIndex = 0
    
    Application.ScreenUpdating = False
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If IsExportableModule(vbComp) Then
            currentIndex = currentIndex + 1
            
            Dim shouldExport As Boolean
            shouldExport = False
            
            If exportAll Then
                shouldExport = True
            Else
                Dim j As Long
                For j = LBound(selectedNumbers) To UBound(selectedNumbers)
                    If Val(Trim(selectedNumbers(j))) = currentIndex Then
                        shouldExport = True
                        Exit For
                    End If
                Next j
            End If
            
            If shouldExport Then
                Dim filePath As String
                filePath = targetFolder & vbComp.name & GetModuleExtension(vbComp)
                
                vbComp.Export filePath
                exportedCount = exportedCount + 1
            End If
        End If
    Next vbComp
    
    Application.ScreenUpdating = True
    
    MsgBox "Export abgeschlossen:" & vbCrLf & vbCrLf & _
           "Exportiert: " & exportedCount & " Modul(e)" & vbCrLf & _
           "Zielordner: " & targetFolder, _
           vbInformation, "Export-Ergebnis"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Exportieren: " & Err.Description, vbCritical, "Export-Fehler"
End Sub

' ============================================================
' UTILITY FUNCTIONS
' ============================================================

' Check if VBA project access is enabled
Private Function HasVBProjectAccess() As Boolean
    On Error GoTo NoAccess
    
    Dim testAccess As Long
    testAccess = ThisWorkbook.VBProject.VBComponents.Count
    
    HasVBProjectAccess = True
    Exit Function
    
NoAccess:
    HasVBProjectAccess = False
End Function

' Check if module is exportable (.bas or .cls, not ThisWorkbook/Sheet)
Private Function IsExportableModule(ByVal vbComp As Object) As Boolean
    Select Case vbComp.Type
        Case vbext_ct_StdModule   ' Standard module (.bas)
            IsExportableModule = True
        Case vbext_ct_ClassModule ' Class module (.cls)
            IsExportableModule = True
        Case Else
            ' Document modules (ThisWorkbook, Sheets) and Forms - not exported
            IsExportableModule = False
    End Select
End Function

' Get file extension for module type
Private Function GetModuleExtension(ByVal vbComp As Object) As String
    Select Case vbComp.Type
        Case vbext_ct_StdModule
            GetModuleExtension = ".bas"
        Case vbext_ct_ClassModule
            GetModuleExtension = ".cls"
        Case vbext_ct_MSForm
            GetModuleExtension = ".frm"
        Case Else
            GetModuleExtension = ".bas"
    End Select
End Function

' Get human-readable module type name
Private Function GetModuleTypeName(ByVal vbComp As Object) As String
    Select Case vbComp.Type
        Case vbext_ct_StdModule
            GetModuleTypeName = "Standard"
        Case vbext_ct_ClassModule
            GetModuleTypeName = "Klasse"
        Case vbext_ct_MSForm
            GetModuleTypeName = "UserForm"
        Case vbext_ct_Document
            GetModuleTypeName = "Dokument"
        Case Else
            GetModuleTypeName = "Unbekannt"
    End Select
End Function

' Get filename from path
Private Function GetFileName(ByVal filePath As String) As String
    Dim pos As Long
    pos = InStrRev(filePath, "\")
    
    If pos > 0 Then
        GetFileName = mid(filePath, pos + 1)
    Else
        GetFileName = filePath
    End If
End Function

' Get filename without extension
Private Function GetFileNameWithoutExtension(ByVal filePath As String) As String
    Dim fileName As String
    fileName = GetFileName(filePath)
    
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    
    If pos > 0 Then
        GetFileNameWithoutExtension = Left(fileName, pos - 1)
    Else
        GetFileNameWithoutExtension = fileName
    End If
End Function

' Browse for folder using FileDialog
Private Function BrowseForFolder(ByVal dialogTitle As String) As String
    BrowseForFolder = ""
    
    Dim fd As fileDialog
    Set fd = Application.fileDialog(msoFileDialogFolderPicker)
    
    With fd
        .title = dialogTitle
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            BrowseForFolder = .SelectedItems(1)
        End If
    End With
End Function

' ============================================================
' QUICK ACCESS MACROS
' ============================================================

' List all modules (for information)
Public Sub ListAllModules()
    On Error GoTo ErrorHandler
    
    If Not HasVBProjectAccess() Then
        MsgBox "Zugriff auf VBA-Projekt nicht moeglich.", vbCritical, "Fehler"
        Exit Sub
    End If
    
    Dim moduleList As String
    moduleList = "Module in dieser Arbeitsmappe:" & vbCrLf & vbCrLf
    
    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        moduleList = moduleList & "  " & vbComp.name & _
                    " (" & GetModuleTypeName(vbComp) & ")" & vbCrLf
    Next vbComp
    
    MsgBox moduleList, vbInformation, "Modulliste"
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler: " & Err.Description, vbCritical, "Fehler"
End Sub
