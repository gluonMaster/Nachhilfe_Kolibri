Attribute VB_Name = "KorrekturStudyload"
Sub RemoveDuplicateEntriesStudyload()
    On Error GoTo ErrorHandler
    
    ' Optimize performance by disabling screen updating and automatic calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Or replace with the specific sheet name, e.g., ThisWorkbook.Sheets("Sheet1")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim key As String
    Dim i As Long
    Dim colsToCompare As Variant
    colsToCompare = Array("A", "B", "C", "D", "F", "G", "H", "AU", "AV")
    
    ' Initialize variables to collect rows for deletion and count deleted rows
    Dim rowsToDelete As Range
    Dim deleteCount As Long
    deleteCount = 0
    
    ' Loop through each row starting from row 11
    For i = 11 To lastRow
        ' Form the key based on the specified columns
        key = ""
        Dim col As Variant
        For Each col In colsToCompare
            key = key & "|" & Trim(CStr(ws.Cells(i, col).value))
        Next col
        
        Dim eValue As String
        eValue = Trim(CStr(ws.Cells(i, "E").value))
        
        If dict.exists(key) Then
            ' Check if there is already a record with column E filled
            If dict(key) = "HasE" Then
                ' If the current record has E empty, mark it for deletion
                If eValue = "" Then
                    If rowsToDelete Is Nothing Then
                        Set rowsToDelete = ws.Rows(i)
                    Else
                        Set rowsToDelete = Union(rowsToDelete, ws.Rows(i))
                    End If
                    deleteCount = deleteCount + 1
                End If
            Else
                ' The previous record did not have E filled
                If eValue <> "" Then
                    ' The current record has E filled, so mark the previous one for deletion
                    Dim j As Long
                    j = dict(key & "_Row")
                    If j > 0 Then
                        If rowsToDelete Is Nothing Then
                            Set rowsToDelete = ws.Rows(j)
                        Else
                            Set rowsToDelete = Union(rowsToDelete, ws.Rows(j))
                        End If
                        deleteCount = deleteCount + 1
                    End If
                    ' Update the dictionary with the current row as having E filled
                    dict(key) = "HasE"
                    dict(key & "_Row") = i
                Else
                    ' Both records have E empty, mark the current one for deletion
                    If rowsToDelete Is Nothing Then
                        Set rowsToDelete = ws.Rows(i)
                    Else
                        Set rowsToDelete = Union(rowsToDelete, ws.Rows(i))
                    End If
                    deleteCount = deleteCount + 1
                End If
            End If
        Else
            ' New unique record
            If eValue <> "" Then
                dict.Add key, "HasE"
            Else
                dict.Add key, "NoE"
            End If
            dict.Add key & "_Row", i
        End If
    Next i
    
    ' Delete all marked rows
    If Not rowsToDelete Is Nothing Then
        rowsToDelete.Delete
        MsgBox "Deleted " & deleteCount & " duplicate records.", vbInformation
    Else
        MsgBox "No duplicate records found.", vbInformation
    End If
    
    ' Restore application settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

