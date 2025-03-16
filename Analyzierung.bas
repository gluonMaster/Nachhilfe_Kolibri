Attribute VB_Name = "Analyzierung"
Sub AnalyzeStudentData()

    'Improve performance by disabling screen updating and automatic calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Define the range of data
    Dim dataRange As Range
    Set dataRange = ws.Range("A5:V" & lastRow)
    
    ' Clear previous formatting and comments
    dataRange.Interior.ColorIndex = xlNone
    dataRange.Columns("V").ClearContents
    Dim cell As Range
    For Each cell In dataRange
        cell.ClearComments
    Next cell
    
    ' Step 1: Check for duplicate records
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim row As Long
    For row = 5 To lastRow
        Dim key As String
        key = ""
        Dim col As Long
        For col = 2 To 15 ' Columns B to O
            key = key & Trim(UCase(ws.Cells(row, col).value)) & "|"
        Next col
        If dict.exists(key) Then
            dict(key) = dict(key) & "," & row
        Else
            dict.Add key, CStr(row)
        End If
    Next row
    
    Dim duplicateRows As Object
    Set duplicateRows = CreateObject("Scripting.Dictionary")
    
    Dim k As Variant
    For Each k In dict.keys
        If InStr(dict(k), ",") > 0 Then
            duplicateRows.Add k, Split(dict(k), ",")
        End If
    Next k
    
    ' Highlight duplicates
    For Each k In duplicateRows.keys
        Dim dupArray() As String
        dupArray = duplicateRows(k)
        Dim i As Long
        Dim commentText As String
        commentText = "Duplicate rows: "
        For i = LBound(dupArray) To UBound(dupArray)
            ws.Range("A" & dupArray(i) & ":V" & dupArray(i)).Interior.Color = vbYellow
            commentText = commentText & dupArray(i) & ", "
        Next i
        ' Remove trailing comma and space
        commentText = Left(commentText, Len(commentText) - 2)
        For i = LBound(dupArray) To UBound(dupArray)
            ws.Cells(CLng(dupArray(i)), "V").value = "There are identical records in rows (" & commentText & ")"
        Next i
    Next k
    
    ' Collect rows to exclude from further checks
    Dim excludedRows As Object
    Set excludedRows = CreateObject("Scripting.Dictionary")
    For Each k In duplicateRows.keys
        Dim dupArr() As String
        dupArr = duplicateRows(k)
        Dim j As Long
        For j = LBound(dupArr) To UBound(dupArr)
            excludedRows.Add CLng(dupArr(j)), True
        Next j
    Next k
    
    ' Step 2: Check for mandatory fields
    Dim mandatoryCols As Variant
    mandatoryCols = Array(2, 3, 4, 5, 7, 8, 12, 13, 15, 19, 20) ' B, C, D, E, G, H, L, M, N, O, S, T
    
    Dim colVariant As Variant ' use a separate variable for iteration
    
    For row = 5 To lastRow
        If Not excludedRows.exists(row) Then
            Dim missingCols As String
            missingCols = ""
            For Each colVariant In mandatoryCols
                If Trim(ws.Cells(row, colVariant).value) = "" Then
                    missingCols = missingCols & ColumnLetter(colVariant) & ","
                    ws.Cells(row, colVariant).Interior.Color = RGB(173, 216, 230) ' Light Blue
                End If
            Next colVariant
            If missingCols <> "" Then
                ' Remove trailing comma
                missingCols = Left(missingCols, Len(missingCols) - 1)
                ws.Cells(row, "V").value = "Missing mandatory information in cells: " & missingCols
            End If
        End If
    Next row
    
    ' Collect rows with missing mandatory fields to exclude from step 3
    Dim excludedRowsStep2 As Object
    Set excludedRowsStep2 = CreateObject("Scripting.Dictionary")
    For row = 5 To lastRow
        If Not excludedRows.exists(row) Then
            For Each colVariant In mandatoryCols
                If Trim(ws.Cells(row, colVariant).value) = "" Then
                    excludedRowsStep2.Add row, True
                    Exit For
                End If
            Next colVariant
        End If
    Next row
    
    ' Step 3: Check consistency for same first and last names
    Dim nameDict As Object
    Set nameDict = CreateObject("Scripting.Dictionary")
    
    For row = 5 To lastRow
        If Not excludedRows.exists(row) And Not excludedRowsStep2.exists(row) Then
            Dim firstName As String
            Dim lastName As String
            firstName = Trim(UCase(ws.Cells(row, "D").value))
            lastName = Trim(UCase(ws.Cells(row, "C").value))
            Dim nameKey As String
            nameKey = lastName & "|" & firstName
            If nameDict.exists(nameKey) Then
                nameDict(nameKey) = nameDict(nameKey) & "," & row
            Else
                nameDict.Add nameKey, CStr(row)
            End If
        End If
    Next row
    
    Dim nameK As Variant
    For Each nameK In nameDict.keys
        If InStr(nameDict(nameK), ",") > 0 Then
            Dim nameArr() As String
            nameArr = Split(nameDict(nameK), ",")
            ' Compare fields B, L, M, N, O
            Dim baseRow As Long
            baseRow = CLng(nameArr(0))
            For i = 1 To UBound(nameArr)
                Dim currentRow As Long
                currentRow = CLng(nameArr(i))
                Dim discrepancies As String
                discrepancies = ""
                
                Dim checkCols As Variant
                checkCols = Array(2, 12, 13, 14, 15) ' B, L, M, N, O
                Dim idx As Variant
                For Each idx In checkCols
                    Dim baseVal As String
                    Dim currentVal As String
                    baseVal = Trim(UCase(ws.Cells(baseRow, idx).value))
                    currentVal = Trim(UCase(ws.Cells(currentRow, idx).value))
                    If baseVal <> currentVal Then
                        discrepancies = discrepancies & ColumnLetter(idx) & ","
                        ws.Cells(currentRow, idx).Interior.Color = RGB(255, 192, 203) ' Pink
                        ws.Cells(baseRow, idx).Interior.Color = RGB(255, 192, 203)
                    End If
                Next idx
                
                If discrepancies <> "" Then
                    ' Remove trailing comma
                    discrepancies = Left(discrepancies, Len(discrepancies) - 1)
                    ws.Cells(currentRow, "V").value = "Possible error in cells: " & discrepancies
                    ws.Cells(baseRow, "V").value = "Possible error in cells: " & discrepancies
                End If
            Next i
        End If
    Next nameK
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Analyse finished.", vbInformation
End Sub

' Function that obtain ColumnLetter from the number of columt (ByVal)
Function ColumnLetter(ByVal colNum As Long) As String
    Dim temp As Long
    Dim letter As String
    temp = colNum
    letter = ""
    Do While temp > 0
        temp = temp - 1
        letter = Chr((temp Mod 26) + 65) & letter
        temp = Int(temp / 26)
    Loop
    ColumnLetter = letter
End Function

