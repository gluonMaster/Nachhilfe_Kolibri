Attribute VB_Name = "CorrectKinderIdentifiers"
Option Explicit

Sub CorrectKinderIdentifiers()
    Dim targetWs As Worksheet
    Dim karteiWs As Worksheet
    Dim externalWb As Workbook
    Dim externalFilePath As String
    Dim lastRowKinder As Long
    Dim lastRowKartei As Long
    Dim rowKinder As Long
    Dim rowKartei As Long
    Dim nameKinder As String
    Dim nameKartei As String
    Dim identifierKinder As String
    Dim identifierKartei As String
    Dim matchFound As Boolean
    Dim errorMessages As String
    Dim dictKartei As Object ' Dictionary to hold name -> identifier mapping
    Dim fNameKinder As String
    Dim lNameKinder As String
    Dim fullNameKinder As String
    Dim parsedNames() As String
    Dim nameKey As String
    Dim prompt As String
    Dim title As String
    Dim dlgButton As Integer
    
    ' Initialize variables
    Set targetWs = ActiveSheet ' Assumes user is on the 'Kinder' sheet
    errorMessages = ""
    
    ' Create a dictionary for Kartei data
    Set dictKartei = CreateObject("Scripting.Dictionary")
    
    ' Prompt user to select the external Kartei workbook
    With Application.FileDialog(msoFileDialogOpen)
        .title = "Select the External Kartei Workbook"
        .Filters.Clear
        .Filters.Add "Excel Workbooks", "*.xls; *.xlsx; *.xlsm"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "No file selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
        externalFilePath = .SelectedItems(1)
    End With
    
    ' Open the external workbook
    On Error GoTo FileOpenError
    Set externalWb = Workbooks.Open(fileName:=externalFilePath, ReadOnly:=True)
    On Error GoTo 0
    
    ' Check if 'Kartei' sheet exists
    On Error Resume Next
    Set karteiWs = externalWb.Sheets("Kartei")
    On Error GoTo 0
    If karteiWs Is Nothing Then
        MsgBox "Sheet 'Kartei' not found in the selected workbook.", vbCritical
        GoTo Cleanup
    End If
    
    ' Determine the last row in Kartei sheet
    lastRowKartei = karteiWs.Cells(karteiWs.Rows.Count, "A").End(xlUp).row
    
    ' Populate the dictionary with Kartei data (Name -> Identifier)
    ' Assumes Names are in column D and Identifiers in column A
    For rowKartei = 2 To lastRowKartei ' Assuming headers in row 1
        nameKartei = Trim(karteiWs.Cells(rowKartei, "D").value)
        identifierKartei = Trim(karteiWs.Cells(rowKartei, "A").value)
        
        ' Normalize the name by replacing delimiters with a single space
        nameKartei = Replace(nameKartei, ";", " ")
        nameKartei = Replace(nameKartei, ",", " ")
        nameKartei = Replace(nameKartei, ".", " ")
        nameKartei = Application.WorksheetFunction.Trim(nameKartei)
        nameKartei = UCase(nameKartei) ' Convert to uppercase for case-insensitive matching
        
        ' Add to dictionary if not already present
        If Not dictKartei.exists(nameKartei) Then
            dictKartei.Add nameKartei, identifierKartei
        Else
            ' Handle duplicate names in Kartei
            ' You can choose to log this or overwrite; here we overwrite
            dictKartei(nameKartei) = identifierKartei
        End If
    Next rowKartei
    
    ' Determine the last row in Kinder sheet
    lastRowKinder = targetWs.Cells(targetWs.Rows.Count, "B").End(xlUp).row ' Assuming identifiers in column B
    
    ' Loop through each child in Kinder sheet starting from row 5
    For rowKinder = 5 To lastRowKinder
        lNameKinder = Trim(targetWs.Cells(rowKinder, "C").value) ' Last name in column C
        fNameKinder = Trim(targetWs.Cells(rowKinder, "D").value) ' First name in column D
        identifierKinder = Trim(targetWs.Cells(rowKinder, "B").value) ' Identifier in column B
        
        ' Skip if either name is missing
        If lNameKinder = "" Or fNameKinder = "" Then
            errorMessages = errorMessages & "Row " & rowKinder & ": Missing last name or first name." & vbCrLf
            GoTo NextChild
        End If
        
        ' Combine last name and first name
        fullNameKinder = lNameKinder & " " & fNameKinder
        fullNameKinder = Application.WorksheetFunction.Trim(fullNameKinder)
        fullNameKinder = UCase(fullNameKinder) ' Convert to uppercase for case-insensitive matching
        
        ' Attempt to find the identifier in the dictionary
        If dictKartei.exists(fullNameKinder) Then
            identifierKartei = dictKartei(fullNameKinder)
            ' Compare identifiers
            If identifierKinder <> identifierKartei Then
                ' Update the identifier in Kinder sheet
                targetWs.Cells(rowKinder, "B").value = identifierKartei
                ' Log the change
                errorMessages = errorMessages & "Row " & rowKinder & ": Identifier corrected from '" & identifierKinder & "' to '" & identifierKartei & "'." & vbCrLf
            End If
        Else
            ' If no match found, log the error
            errorMessages = errorMessages & "Row " & rowKinder & ": No matching name found in Kartei." & vbCrLf
        End If
        
NextChild:
    Next rowKinder
    
    ' Inform the user about the results
    If errorMessages <> "" Then
        ' Display the errors and changes
        ' Optionally, you can write these to a log sheet instead
        MsgBox "Identifier correction completed with the following details:" & vbCrLf & errorMessages, vbInformation
    Else
        MsgBox "Identifier correction completed successfully. No discrepancies found.", vbInformation
    End If
    
Cleanup:
    ' Close the external workbook without saving
    If Not externalWb Is Nothing Then
        externalWb.Close SaveChanges:=False
    End If
    
    ' Release objects
    Set dictKartei = Nothing
    Set karteiWs = Nothing
    Set externalWb = Nothing
    Set targetWs = Nothing
    
    Exit Sub
    
FileOpenError:
    MsgBox "Unable to open the selected file. Please ensure it's a valid Excel workbook.", vbCritical
    Resume Cleanup
End Sub

