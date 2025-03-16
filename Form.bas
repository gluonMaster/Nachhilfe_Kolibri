Attribute VB_Name = "Form"
Sub FillFormFromActiveCell()
    ' This procedure fills the Form sheet with data from the active cell's row
    ' It copies specific cell values and formatting from the active sheet to the Form sheet
    
    Dim sourceSheet As Worksheet
    Dim formSheet As Worksheet
    Dim activeRow As Long
    Dim sourceCell As Range
    Dim targetCell As Range
    Dim sourceRange As Range
    Dim targetRange As Range
    
    ' Store current active sheet and row
    Set sourceSheet = ActiveSheet
    activeRow = ActiveCell.row
    
    ' Reference the Form sheet
    On Error Resume Next
    Set formSheet = ThisWorkbook.Sheets("Form")
    On Error GoTo 0
    
    ' Check if Form sheet exists
    If formSheet Is Nothing Then
        MsgBox "Sheet 'Form' not found!", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 1) Put active row number in cell A1 of Form sheet
    formSheet.Range("A1").value = activeRow
    
    ' 2) Put active sheet name in cell A2 of Form sheet
    formSheet.Range("C1").value = sourceSheet.name
    
    ' 3) Copy values from cells A, B, C to cells A5, B5, C5 on Form sheet
    formSheet.Range("A5").value = sourceSheet.Cells(activeRow, "A").value
    formSheet.Range("B5").value = sourceSheet.Cells(activeRow, "B").value
    formSheet.Range("C5").value = sourceSheet.Cells(activeRow, "C").value
    
    ' 4) Copy values from cells D, E, F, G to cells C7, C9, B9, B10 on Form sheet
    formSheet.Range("C7").value = sourceSheet.Cells(activeRow, "D").value
    formSheet.Range("C9").value = sourceSheet.Cells(activeRow, "E").value
    formSheet.Range("B9").value = sourceSheet.Cells(activeRow, "F").value
    formSheet.Range("B10").value = sourceSheet.Cells(activeRow, "G").value
    
    ' 5) Copy values and formatting from ranges to Form sheet
    ' Define source and target ranges
    Dim rangePairs As Variant
    rangePairs = Array( _
        Array(sourceSheet.Range("J" & activeRow & ":P" & activeRow), formSheet.Range("D4:J4")), _
        Array(sourceSheet.Range("Q" & activeRow & ":W" & activeRow), formSheet.Range("D6:J6")), _
        Array(sourceSheet.Range("X" & activeRow & ":AD" & activeRow), formSheet.Range("D8:J8")), _
        Array(sourceSheet.Range("AE" & activeRow & ":AK" & activeRow), formSheet.Range("D10:J10")), _
        Array(sourceSheet.Range("AL" & activeRow & ":AN" & activeRow), formSheet.Range("D12:F12")) _
    )
    
    ' Process each range pair
    Dim i As Integer
    For i = 0 To UBound(rangePairs)
        Set sourceRange = rangePairs(i)(0)
        Set targetRange = rangePairs(i)(1)
        
        ' Copy values
        targetRange.value = sourceRange.value
        
        ' Copy formatting (interior color)
        Dim j As Integer
        For j = 1 To sourceRange.Cells.Count
            targetRange.Cells(j).Interior.Color = sourceRange.Cells(j).Interior.Color
        Next j
    Next i
    
    Application.ScreenUpdating = True
    
    ' Activate the Form sheet
    formSheet.Activate
    
    'MsgBox "Form filled successfully!", vbInformation
End Sub

Sub TransferDataFromForm()
    ' This procedure transfers data from Form sheet back to the original sheet
    ' and clears the Form sheet afterward
    
    Dim formSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim targetSheetName As String
    Dim sourceRange As Range
    Dim targetRange As Range
    
    ' Reference the Form sheet
    On Error Resume Next
    Set formSheet = ThisWorkbook.Sheets("Form")
    On Error GoTo 0
    
    If formSheet.Range("C9") = "" Then
        MsgBox "You did not specify the type of the lesson. Data transfer has not been performed!", vbInformation
        Exit Sub
    End If
    
    ' Check if Form sheet exists
    If formSheet Is Nothing Then
        MsgBox "Sheet 'Form' not found!", vbExclamation
        Exit Sub
    End If
    
    ' Get target row number from cell A1
    On Error Resume Next
    targetRow = CLng(formSheet.Range("A1").value)
    On Error GoTo 0
    
    If targetRow <= 0 Then
        MsgBox "Invalid row number in cell A1!", vbExclamation
        Exit Sub
    End If
    
    ' Get target sheet name from cell C1
    targetSheetName = formSheet.Range("C1").value
    
    ' Check if target sheet exists
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Sheets(targetSheetName)
    On Error GoTo 0
    
    If targetSheet Is Nothing Then
        MsgBox "Target sheet '" & targetSheetName & "' not found!", vbExclamation
        Exit Sub
    End If
    
    If ClearAllFilters(targetSheet) Then
        Application.ScreenUpdating = False
        
        ' Transfer data from Form sheet to target sheet
        
        ' Define range pairs (source range on Form, target range on target sheet)
        Dim rangePairs As Variant
        rangePairs = Array( _
            Array(formSheet.Range("D4:J4"), targetSheet.Range("J" & targetRow & ":P" & targetRow)), _
            Array(formSheet.Range("D6:J6"), targetSheet.Range("Q" & targetRow & ":W" & targetRow)), _
            Array(formSheet.Range("D8:J8"), targetSheet.Range("X" & targetRow & ":AD" & targetRow)), _
            Array(formSheet.Range("D10:J10"), targetSheet.Range("AE" & targetRow & ":AK" & targetRow)), _
            Array(formSheet.Range("D12:F12"), targetSheet.Range("AL" & targetRow & ":AN" & targetRow)) _
        )
        
        ' Transfer value from C9 to E in the target row
        targetSheet.Cells(targetRow, "E").value = formSheet.Range("C9").value
        
        ' Process each range pair
        Dim i As Integer
        For i = 0 To UBound(rangePairs)
            Set sourceRange = rangePairs(i)(0)
            Set targetRange = rangePairs(i)(1)
            
            ' Copy values
            targetRange.value = sourceRange.value
            
            ' Copy formatting (interior color)
    '        Dim j As Integer
    '        For j = 1 To sourceRange.Cells.Count
    '            targetRange.Cells(j).Interior.Color = sourceRange.Cells(j).Interior.Color
    '        Next j
        Next i
        
        ' Clear the Form sheet
        
        ' Clear specific cells
        Dim cellsToClear As Variant
        cellsToClear = Array("A1", "C1", "A5", "B5", "C5", "C7", "C9", "B9", "B10")
        
        For i = 0 To UBound(cellsToClear)
            formSheet.Range(cellsToClear(i)).ClearContents
        Next i
        
        ' Clear ranges including their formatting
        Dim rangesToClear As Variant
        rangesToClear = Array("D4:J4", "D6:J6", "D8:J8", "D10:J10", "D12:F12")
        
        For i = 0 To UBound(rangesToClear)
            formSheet.Range(rangesToClear(i)).ClearContents
            formSheet.Range(rangesToClear(i)).Interior.ColorIndex = xlNone
        Next i
        
        ' Activate the target sheet
        targetSheet.Activate
        
        Application.ScreenUpdating = True
    End If
    
    'MsgBox "Data transferred successfully!", vbInformation
End Sub

