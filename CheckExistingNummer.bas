Attribute VB_Name = "CheckExistingNummer"
Option Explicit

Public Sub CheckAndFixExistingNumbers()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim cell As Range
    Dim invalidCount As Long
    Dim originalValue As String
    Dim correctedValue As String

    ' Set the worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    'Improve performance by disabling screen updating and automatic calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Find the last used row in column B
    If Application.WorksheetFunction.CountA(ws.Columns("B")) < 5 Then
        ' No data starting from row 5
        Exit Sub
    Else
        lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    End If
    
    invalidCount = 0
    
    ' Loop through rows starting from 5
    For currentRow = 5 To lastRow
        Set cell = ws.Cells(currentRow, "B")
        
        If Not IsEmpty(cell) Then
            originalValue = Trim(CStr(cell.value))
            
            ' Replace commas with dots
            originalValue = Replace(originalValue, ",", ".")
            
            ' Trim spaces
            originalValue = Application.WorksheetFunction.Trim(originalValue)
            
            ' Check if format is valid
            If Not IsValidFormat(originalValue) Then
                correctedValue = CorrectFormat(originalValue)
                If correctedValue <> "" Then
                    ' If corrected successfully
                    cell.NumberFormat = "@"
                    cell.value = correctedValue
                Else
                    ' If cannot correct, highlight the cell
                    cell.Interior.Color = vbYellow
                    invalidCount = invalidCount + 1
                End If
            Else
                ' If format is valid but we still ensure the cell is text format
                cell.NumberFormat = "@"
                cell.value = originalValue
            End If
        End If
    Next currentRow
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' If there are invalid entries that could not be corrected, show a message
    If invalidCount > 0 Then
        MsgBox "There are " & invalidCount & " invalid entries in column B that could not be corrected automatically. They have been highlighted in yellow.", vbExclamation, "Invalid Format"
    End If
End Sub

' Ensure that these functions are in the same module or accessible from this module

Private Function IsValidFormat(ByVal textVal As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Pattern = "^\d\.\s\d{4}$"
        .ignoreCase = True
        .Global = False
    End With
    
    IsValidFormat = regex.Test(textVal)
End Function

Private Function CorrectFormat(ByVal textVal As String) As String
    Dim parts() As String
    Dim firstPart As String, secondPart As String

    ' Try splitting by dot
    parts = Split(textVal, ".")
    If UBound(parts) = 1 Then
        firstPart = Trim(parts(0))
        secondPart = Trim(parts(1))
        If IsNumeric(firstPart) And IsNumeric(secondPart) Then
            If Len(firstPart) = 1 And Len(secondPart) = 4 Then
                CorrectFormat = firstPart & ". " & secondPart
                Exit Function
            End If
        End If
    End If
    
    ' Try splitting by comma
    parts = Split(textVal, ",")
    If UBound(parts) = 1 Then
        firstPart = Trim(parts(0))
        secondPart = Trim(parts(1))
        If IsNumeric(firstPart) And IsNumeric(secondPart) Then
            If Len(firstPart) = 1 And Len(secondPart) = 4 Then
                CorrectFormat = firstPart & ". " & secondPart
                Exit Function
            End If
        End If
    End If
    
    ' If no correction possible
    CorrectFormat = ""
End Function

