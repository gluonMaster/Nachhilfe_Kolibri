Attribute VB_Name = "RemoveDuplicates"
Option Explicit

Sub RemoveAndReportDuplicates()
    Dim wsKinder As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim key As Variant
    Dim dictDuplicates As Object
    Dim entry As Variant
    Dim rowsToDelete As Collection
    Dim conflictingRows As String
    Dim identifiers As Object
    Dim totalRows As Long
    Dim processedRows As Long
    Dim i As Long
    
    ' Initialize
    Set wsKinder = ThisWorkbook.ActiveSheet ' Assumes the user is on the 'Kinder' sheet
    Set dictDuplicates = CreateObject("Scripting.Dictionary")
    Set rowsToDelete = New Collection
    conflictingRows = ""
    
    ' Determine the last row with data in column C (assuming column C always has data if the row is used)
    lastRow = wsKinder.Cells(wsKinder.Rows.Count, "C").End(xlUp).row
    totalRows = lastRow - 4 ' Since data starts from row 5
    
    ' Initialize the Progress Form
    With frmProgress
        .cancelRequested = False
        .lblProgress.Width = 0
        .lblStatus.Caption = "Progress: 0%"
        .Show vbModeless ' Show the form modelessly so the macro can continue running
    End With
    
    ' Improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Loop through each row starting from row 5
    For row = 5 To lastRow
        ' Update the progress bar
        processedRows = row - 4
        With frmProgress
            .lblProgress.Width = (.fraProgress.Width) * (processedRows / totalRows)
            .lblStatus.Caption = "Progress: " & Format((processedRows / totalRows) * 100, "0") & "%"
        End With
        DoEvents ' Allow the form to update
        
        ' Check if cancellation was requested
        If frmProgress.cancelRequested Then
            MsgBox "Operation cancelled by the user.", vbExclamation, "Cancelled"
            GoTo Cleanup
        End If
        
        ' Build the key by concatenating the values of columns C, D, E, F, G, H, I, J, L, O
        key = Trim(wsKinder.Cells(row, "C").value) & "|" & _
              Trim(wsKinder.Cells(row, "D").value) & "|" & _
              Trim(wsKinder.Cells(row, "E").value) & "|" & _
              Trim(wsKinder.Cells(row, "F").value) & "|" & _
              Trim(wsKinder.Cells(row, "G").value) & "|" & _
              Trim(wsKinder.Cells(row, "H").value) & "|" & _
              Trim(wsKinder.Cells(row, "I").value) & "|" & _
              Trim(wsKinder.Cells(row, "J").value) & "|" & _
              Trim(wsKinder.Cells(row, "L").value) & "|" & _
              Trim(wsKinder.Cells(row, "O").value)
        
        ' Check if the key already exists in the dictionary
        If dictDuplicates.exists(key) Then
            ' If the key exists, append the current row and identifier to the existing entry
            dictDuplicates(key).Add Array(row, wsKinder.Cells(row, "B").value)
        Else
            ' If the key does not exist, create a new entry with the current row and identifier
            Set dictDuplicates(key) = New Collection
            dictDuplicates(key).Add Array(row, wsKinder.Cells(row, "B").value)
        End If
    Next row
    
    ' Iterate through the dictionary to identify duplicates
    For Each key In dictDuplicates.keys
        If dictDuplicates(key).Count > 1 Then
            ' Initialize a dictionary to track unique identifiers within this group
            Set identifiers = CreateObject("Scripting.Dictionary")
            
            ' Populate the identifiers dictionary
            For Each entry In dictDuplicates(key)
                identifiers(Trim(entry(1))) = True
            Next entry
            
            If identifiers.Count = 1 Then
                ' All identifiers are the same; mark all but the first row for deletion
                For i = 2 To dictDuplicates(key).Count
                    rowsToDelete.Add dictDuplicates(key)(i)(0)
                Next i
            Else
                ' Identifiers differ; collect the row numbers for reporting
                For Each entry In dictDuplicates(key)
                    conflictingRows = conflictingRows & "Row " & entry(0) & vbCrLf
                Next entry
            End If
        End If
    Next key
    
    ' Delete the duplicate rows (from bottom to top to avoid shifting issues)
    If rowsToDelete.Count > 0 Then
        Dim arrDelete() As Long
        ReDim arrDelete(1 To rowsToDelete.Count)
        For i = 1 To rowsToDelete.Count
            arrDelete(i) = rowsToDelete(i)
        Next i
        ' Sort the array in descending order
        Call QuickSort(arrDelete, LBound(arrDelete), UBound(arrDelete))
        
        ' Delete the rows
        For i = 1 To UBound(arrDelete)
            wsKinder.Rows(arrDelete(i)).Delete
        Next i
    End If
    
    ' Close the Progress Form
    Unload frmProgress
    
    ' Restore Excel settings
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Display conflicting rows if any
    If conflictingRows <> "" Then
        MsgBox "Conflicting duplicate records found in the following rows (identifiers differ):" & vbCrLf & conflictingRows, vbExclamation, "Duplicate Identifiers Found"
    Else
        MsgBox "Duplicate records with identical identifiers have been removed successfully.", vbInformation, "Operation Completed"
    End If
    
    ' Cleanup
    Set dictDuplicates = Nothing
    Set rowsToDelete = Nothing
    Set identifiers = Nothing
    
End Sub

' QuickSort algorithm to sort an array in descending order
Sub QuickSort(arr() As Long, first As Long, last As Long)
    Dim pivot As Long
    Dim i As Long, j As Long
    Dim temp As Long
    
    If first < last Then
        pivot = first
        i = first
        j = last
        
        While i < j
            While arr(i) >= arr(pivot) And i < last
                i = i + 1
            Wend
            While arr(j) < arr(pivot)
                j = j - 1
            Wend
            If i < j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Wend
        
        temp = arr(pivot)
        arr(pivot) = arr(j)
        arr(j) = temp
        
        Call QuickSort(arr, first, j - 1)
        Call QuickSort(arr, j + 1, last)
    End If
End Sub


