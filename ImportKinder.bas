Attribute VB_Name = "ImportKinder"
Option Explicit

Sub ImportDataFromSource()
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim sourceFilePath As String
    Dim lastRow As Long
    Dim targetLastRow As Long
    Dim i As Long
    Dim surname As String
    Dim name As String
    Dim fullName As String
    Dim arrNames() As String
    Dim fd As FileDialog
    Dim userSelected As Integer
    
    ' **1. Prompt user to select the source workbook**
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .title = "Select the Source Workbook (Nachhilfe_Ubersichet.xlsx)"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xlsb; *.xls"
        .AllowMultiSelect = False
        userSelected = .Show
        If userSelected = -1 Then
            sourceFilePath = .SelectedItems(1)
        Else
            MsgBox "No file selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' **2. Open the source workbook**
    On Error Resume Next
    Set sourceWb = Workbooks.Open(sourceFilePath)
    If sourceWb Is Nothing Then
        MsgBox "Source workbook could not be opened.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' **3. Set the source and target worksheets**
    On Error Resume Next
    Set sourceWs = sourceWb.Sheets("Kartei")
    If sourceWs Is Nothing Then
        MsgBox "Sheet 'Kartei' not found in the source workbook.", vbExclamation
        sourceWb.Close SaveChanges:=False
        Exit Sub
    End If
    On Error GoTo 0
    
    Set targetWs = ThisWorkbook.Sheets("Kinder_pre")
    If targetWs Is Nothing Then
        MsgBox "Sheet 'Kinder_pre' not found in the target workbook.", vbExclamation
        sourceWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' **4. Find the last row in source sheet**
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, "A").End(xlUp).row
    
    ' **5. Find the last row in target sheet**
    targetLastRow = targetWs.Cells(targetWs.Rows.Count, "A").End(xlUp).row
    If targetLastRow < 1 Then targetLastRow = 1
    
    ' **6. Clear existing data except headers in target sheet**
    targetWs.Range("A3:G" & targetWs.Rows.Count).ClearContents
    targetLastRow = 2
    
    ' **7. Improve performance**
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' **8. Loop through source rows**
    For i = 2 To lastRow
        If Trim(sourceWs.Cells(i, "C").value) <> "" Then
            ' **a. Read data from source**
            Dim childID As String
            Dim boundaryDate As Variant
            Dim fullNameRaw As String
            Dim birthDate As Variant
            Dim address As String
            Dim subjects As String
            
            childID = sourceWs.Cells(i, "A").value
            boundaryDate = sourceWs.Cells(i, "C").value
            fullNameRaw = sourceWs.Cells(i, "D").value
            birthDate = sourceWs.Cells(i, "E").value
            address = sourceWs.Cells(i, "F").value
            subjects = sourceWs.Cells(i, "J").value
            
            ' **b. Process full name to get surname and name**
            fullName = fullNameRaw
            ' Replace ";" and "," with space
            fullName = Replace(fullName, ";", " ")
            fullName = Replace(fullName, ",", " ")
            ' Replace multiple spaces with single space
            Do While InStr(fullName, "  ") > 0
                fullName = Replace(fullName, "  ", " ")
            Loop
            fullName = Trim(fullName)
            ' Split by space
            arrNames = Split(fullName, " ")
            If UBound(arrNames) >= 1 Then
                surname = arrNames(0)
                ' Join the rest as name
                name = ""
                Dim jIdx As Long
                For jIdx = 1 To UBound(arrNames)
                    name = name & arrNames(jIdx) & " "
                Next jIdx
                name = Trim(name)
            Else
                ' If name could not be split, assign full name to surname, leave name empty
                surname = fullName
                name = ""
            End If
            
            ' **c. Write data to target sheet**
            targetLastRow = targetLastRow + 1
            targetWs.Cells(targetLastRow, "A").value = childID
            targetWs.Cells(targetLastRow, "B").value = surname
            targetWs.Cells(targetLastRow, "C").value = name
            targetWs.Cells(targetLastRow, "D").value = boundaryDate
            targetWs.Cells(targetLastRow, "E").value = birthDate
            targetWs.Cells(targetLastRow, "F").value = address
            targetWs.Cells(targetLastRow, "G").value = subjects
        End If
    Next i
    
    ' **9. Restore settings**
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' **10. Close the source workbook**
    sourceWb.Close SaveChanges:=False
    
    MsgBox "Data import completed successfully.", vbInformation
End Sub

