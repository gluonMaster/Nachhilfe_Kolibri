Attribute VB_Name = "SurenameNameSplitting"
Option Explicit

Sub SplitSurnameName()
    Dim selectedRange As Range
    Dim cell As Range
    Dim fullName As String
    Dim surname As String
    Dim namePart As String
    Dim arrNames() As String
    Dim i As Long
    
    ' **1. Set the selected range**
    On Error Resume Next
    Set selectedRange = Application.Selection
    If selectedRange Is Nothing Then
        MsgBox "No cells are selected. Please select the range to process.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' **2. Improve performance**
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' **3. Loop through each cell in the selected range**
    For Each cell In selectedRange
        ' **a. Check if the cell is not empty and contains text**
        If Not IsEmpty(cell) And IsText(cell.value) Then
            fullName = Trim(cell.value)
            
            ' **b. Replace delimiters ("," and ";") with space**
            fullName = Replace(fullName, ",", " ")
            fullName = Replace(fullName, ";", " ")
            
            ' **c. Replace multiple spaces with a single space**
            Do While InStr(fullName, "  ") > 0
                fullName = Replace(fullName, "  ", " ")
            Loop
            
            fullName = Trim(fullName)
            
            ' **d. Split the full name by space**
            arrNames = Split(fullName, " ")
            
            If UBound(arrNames) >= 1 Then
                ' **e. Assign surname and name**
                surname = arrNames(0)
                namePart = ""
                For i = 1 To UBound(arrNames)
                    namePart = namePart & arrNames(i) & " "
                Next i
                namePart = Trim(namePart)
            Else
                ' **f. If only one part is found, assign it to surname and leave name blank**
                surname = fullName
                namePart = ""
            End If
            
            ' **g. Update the cells**
            cell.value = surname
            cell.Offset(0, 1).value = namePart
        End If
    Next cell
    
    ' **4. Restore settings**
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Surname and name have been successfully split.", vbInformation
End Sub

' **Helper Function to Check if Cell Contains Text**
Function IsText(value As Variant) As Boolean
    IsText = VarType(value) = vbString
End Function

