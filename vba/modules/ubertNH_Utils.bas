Attribute VB_Name = "ubertNH_Utils"
Option Explicit

' ====================================================================
' Module: ubertNH_Utils
' Description: Utility functions for text processing, filters, etc.
' ====================================================================

' Clean text by removing spaces, commas, and semicolons
Public Function CleanText(ByVal inputText As String) As String
    Dim result As String
    result = inputText
    
    ' Remove spaces
    result = Replace(result, " ", "")
    
    ' Remove commas
    result = Replace(result, ",", "")
    
    ' Remove semicolons
    result = Replace(result, ";", "")
    
    ' Convert to lowercase for case-insensitive comparison
    result = LCase(result)
    
    CleanText = result
End Function

' Concatenate first and last name and clean
Public Function GetCleanedFullName(ByVal lastName As String, ByVal firstName As String) As String
    GetCleanedFullName = CleanText(lastName & firstName)
End Function

' Check if subject name contains search term
Public Function SubjectMatches(ByVal targetSubject As String, ByVal sourceSubject As String) As Boolean
    Dim cleanTarget As String
    Dim cleanSource As String
    
    ' Clean both subjects: lowercase and remove ALL spaces (not just trim)
    cleanTarget = LCase(Replace(Trim(targetSubject), " ", ""))
    cleanSource = LCase(Replace(Trim(sourceSubject), " ", ""))
    
    ' Check if target contains source (case-insensitive, space-insensitive)
    SubjectMatches = (InStr(1, cleanTarget, cleanSource, vbTextCompare) > 0)
End Function

' Get semester number based on month
Public Function GetSemester(ByVal monthNum As Integer) As Integer
    Dim semester1Arr() As String
    Dim i As Integer
    
    semester1Arr = Split(ubertNH_Config.SEMESTER1_MONTHS, ",")
    
    For i = LBound(semester1Arr) To UBound(semester1Arr)
        If CInt(semester1Arr(i)) = monthNum Then
            GetSemester = 1
            Exit Function
        End If
    Next i
    
    GetSemester = 2
End Function

' Get column letter for month (U = January, V = February, ..., AF = December)
Public Function GetMonthColumn(ByVal monthNum As Integer) As String
    ' U = 21st column (January = month 1)
    ' V = 22nd column (February = month 2)
    ' ... AF = 32nd column (December = month 12)
    
    Dim colNumber As Integer
    colNumber = 20 + monthNum  ' U is column 21, so 20 + 1 = 21
    
    GetMonthColumn = ConvertToColumnLetter(colNumber)
End Function

' Convert column number to letter
Private Function ConvertToColumnLetter(ByVal colNum As Integer) As String
    Dim result As String
    Dim num As Integer
    
    num = colNum
    Do While num > 0
        Dim remainder As Integer
        remainder = (num - 1) Mod 26
        result = Chr(65 + remainder) & result
        num = (num - remainder) \ 26
    Loop
    
    ConvertToColumnLetter = result
End Function

' Clear filters on a worksheet (remove filter criteria but keep AutoFilter arrows)
Public Sub ClearFilters(ByVal ws As Worksheet)
    On Error Resume Next
    ' Check if AutoFilter is enabled and has filters applied
    If ws.AutoFilterMode Then
        If ws.FilterMode Then
            ' Remove filter criteria but keep AutoFilter arrows
            ws.ShowAllData
        End If
    End If
    On Error GoTo 0
End Sub

' Get last row with data in specified column
Public Function GetLastRow(ByVal ws As Worksheet, ByVal colLetter As String) As Long
    Dim lastRow As Long
    
    With ws
        lastRow = .Cells(.rows.Count, colLetter).End(xlUp).row
    End With
    
    GetLastRow = lastRow
End Function

' Get last row checking two columns (A or B not empty)
Public Function GetLastRowTwoColumns(ByVal ws As Worksheet, ByVal col1 As String, ByVal col2 As String) As Long
    Dim lastRowA As Long
    Dim lastRowB As Long
    
    lastRowA = GetLastRow(ws, col1)
    lastRowB = GetLastRow(ws, col2)
    
    ' Return the maximum
    If lastRowA > lastRowB Then
        GetLastRowTwoColumns = lastRowA
    Else
        GetLastRowTwoColumns = lastRowB
    End If
End Function

' Check if cell is empty or zero
Public Function IsCellEmptyOrZero(ByVal cellValue As Variant) As Boolean
    If IsEmpty(cellValue) Then
        IsCellEmptyOrZero = True
    ElseIf IsNumeric(cellValue) Then
        IsCellEmptyOrZero = (cellValue = 0)
    Else
        IsCellEmptyOrZero = (Trim(cellValue) = "")
    End If
End Function

' Highlight range with pale pink color
Public Sub HighlightRange(ByVal rng As Range)
    rng.Interior.Color = ubertNH_Config.HIGHLIGHT_COLOR_RGB
End Sub

' Format message with placeholders
Public Function FormatMessage(ByVal template As String, ParamArray values() As Variant) As String
    Dim result As String
    Dim i As Integer
    Dim placeholders As Variant
    
    result = template
    
    placeholders = Array("{SUCCESS}", "{NOTFOUND}", "{OVERWRITES}", "{LOGPATH}", "{ERROR}")
    
    For i = LBound(values) To UBound(values)
        If i <= UBound(placeholders) Then
            result = Replace(result, placeholders(i), CStr(values(i)))
        End If
    Next i
    
    FormatMessage = result
End Function

' Get current timestamp for log file name
Public Function GetTimestamp() As String
    GetTimestamp = Format(Now, "DDMMYYYY_HHMMSS")
End Function
