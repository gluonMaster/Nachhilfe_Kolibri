Attribute VB_Name = "ExemplareMachen"
Option Explicit

Sub CreateCopiesForEmployees()
    On Error GoTo ErrorHandler
    
    ' Declare necessary variables
    Dim employees As Variant
    Dim fd As FileDialog
    Dim selectedFolder As String
    Dim originalFileName As String
    Dim fileExtension As String
    Dim newFileName As String
    Dim fullPath As String
    Dim employee As Variant
    Dim fileExists As Boolean
    
    ' Define the list of employee names
    employees = Array("Valya", "Alla")
    
    ' **1. Prompt the user to select the destination folder**
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .title = "Select Destination Folder for File Copies"
        .AllowMultiSelect = False
        If .Show = -1 Then ' If the user selects a folder
            selectedFolder = .SelectedItems(1)
        Else
            ' If the user cancels the dialog, exit the macro
            MsgBox "Operation cancelled by the user.", vbInformation, "Cancelled"
            Exit Sub
        End If
    End With
    
    ' **2. Extract the original file name and extension**
    originalFileName = ThisWorkbook.name
    fileExtension = mid(originalFileName, InStrRev(originalFileName, "."))
    originalFileName = Left(originalFileName, InStrRev(originalFileName, ".") - 1)
    
    ' **3. Check for existing files to prevent overwriting**
    fileExists = False
    For Each employee In employees
        newFileName = originalFileName & "_" & employee & ".xlsm"
        fullPath = selectedFolder & "\" & newFileName
        If Dir(fullPath) <> "" Then
            fileExists = True
            Exit For
        End If
    Next employee
    
    ' **4. If any file exists, display a message and exit**
    If fileExists Then
        MsgBox "Files with the intended names already exist in the selected directory. " & _
               "Please synchronize data from the copies or manually delete the old file copies in the target directory.", _
               vbExclamation, "Duplicate Files Found"
        Exit Sub
    End If
    
    ' **5. Improve performance by disabling screen updating and automatic calculation**
    Application.ScreenUpdating = False
    
    ' **6. Create copies for each employee**
    For Each employee In employees
        newFileName = originalFileName & "_" & employee & ".xlsm"
        fullPath = selectedFolder & "\" & newFileName
        
        ThisWorkbook.SaveCopyAs fileName:=fullPath
        
    Next employee
    
    ' **7. Restore settings**
    Application.ScreenUpdating = True
    
    ' **8. Inform the user of successful completion**
    MsgBox "Copies created successfully for all employees.", vbInformation, "Success"
    Exit Sub

ErrorHandler:
    ' **9. Handle any unexpected errors gracefully**
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    Application.ScreenUpdating = True
End Sub


