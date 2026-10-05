Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ExitHandler

    ' Prevent the event from triggering recursively
    Application.EnableEvents = False

    Dim ws As Worksheet
    Set ws = Me ' Current sheet (e.g., "Kinder")

    Dim wsArchiv As Worksheet
    Set wsArchiv = ThisWorkbook.Sheets("Archiv")

    Dim lastUsedRow As Long
    Dim NextRow As Long
    Dim cell As Range
    Dim serialNumber As Long
    Dim rowsToClear As Collection
    Set rowsToClear = New Collection

    Const START_ROW As Long = 5

    ' Find the last used row in column A (Serial Numbers)
    If Application.WorksheetFunction.CountA(ws.Columns("A")) < START_ROW Then
        lastUsedRow = START_ROW - 1
    Else
        lastUsedRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    End If

    NextRow = lastUsedRow + 1

    ' ---- Existing logic for columns C and D ----
    If Not Intersect(Target, ws.Columns("C:D")) Is Nothing Then
        For Each cell In Intersect(Target, ws.Columns("C:D"))
            Dim currentRow As Long
            currentRow = cell.row

            If currentRow < START_ROW Then GoTo NextCell

            ' If either C or D is not empty
            If Application.WorksheetFunction.CountA(ws.Range("C" & currentRow & ":D" & currentRow)) > 0 Then
                ' Assign serial number if column A is empty
                If ws.Cells(currentRow, "A").value = "" Then
                    If lastUsedRow < START_ROW Then
                        serialNumber = 1
                    Else
                        serialNumber = ws.Cells(lastUsedRow, "A").value + 1
                    End If
                    ws.Cells(currentRow, "A").value = serialNumber
                End If

                ' Assign the DATEDIF formula in column M if it's empty
                If ws.Cells(currentRow, "M").Formula = "" Then
                    ws.Cells(currentRow, "M").FormulaR1C1 = "=DATEDIF(RC[-1], TODAY(), ""Y"")"
                End If

                ' Copy data from Archiv if needed
                Call CopyDataFromArchiv(ws, wsArchiv, currentRow)

            Else
                ' If both C and D are empty, clear the entire row
                If currentRow >= START_ROW Then
                    rowsToClear.Add currentRow
                End If
            End If
NextCell:
        Next cell
    End If

    ' ---- New logic for column B ----
    If Not Intersect(Target, ws.Columns("B")) Is Nothing Then
        Dim changedCell As Range
        For Each changedCell In Intersect(Target, ws.Columns("B"))
            If Not IsEmpty(changedCell.value) Then
                Dim originalValue As String
                originalValue = Trim(changedCell.value)

                ' Set cell format to text to avoid automatic date conversion
                changedCell.numberFormat = "@"

                ' Replace commas with dots
                originalValue = Replace(originalValue, ",", ".")

                ' Trim spaces
                originalValue = Application.WorksheetFunction.Trim(originalValue)

                ' Check if format is valid
                If Not IsValidFormat(originalValue) Then
                    Dim correctedValue As String
                    correctedValue = CorrectFormat(originalValue)

                    If correctedValue <> "" Then
                        ' If correction succeeded
                        changedCell.value = correctedValue
                    Else
                        ' If correction failed, clear the cell and show a message
                        MsgBox "Wrong number format. Please use X. XXXX format.", vbExclamation, "Format Error"
                        changedCell.ClearContents
                    End If
                Else
                    ' If format is valid, just reassign the value
                    changedCell.value = originalValue
                End If
            End If
        Next changedCell
    End If

    ' Clear the collected rows if any
    Dim rowToClear As Variant
    For Each rowToClear In rowsToClear
        ws.rows(rowToClear).ClearContents
    Next rowToClear

ExitHandler:
    ' Re-enable events
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "An error occurred: " & Err.Description, vbExclamation, "Error"
    End If
End Sub

Private Sub CopyDataFromArchiv(wsKinder As Worksheet, wsArchiv As Worksheet, currentRow As Long)
    Dim ignoreCase As Boolean
    Dim ignoreSpaces As Boolean
    Dim searchName As String
    Dim searchSurname As String
    Dim valueC As String
    Dim valueD As String
    Dim isRowEmpty As Boolean
    Dim cellsToCheck As Variant
    Dim cellAddress As Variant
    Dim lastRowArchiv As Long
    Dim foundRow As Long
    Dim compareC As String
    Dim compareD As String
    Dim i As Long

    ignoreCase = True
    ignoreSpaces = True

    cellsToCheck = Array("B", "G", "H", "K", "L", "O")
    isRowEmpty = True
    For Each cellAddress In cellsToCheck
        If wsKinder.Cells(currentRow, cellAddress).value <> "" Then
            isRowEmpty = False
            Exit For
        End If
    Next cellAddress

    If isRowEmpty Then
        valueC = wsKinder.Cells(currentRow, "C").value
        valueD = wsKinder.Cells(currentRow, "D").value

        If Trim(valueC) <> "" And Trim(valueD) <> "" Then
            If ignoreSpaces Then
                searchName = Replace(valueC, " ", "")
                searchSurname = Replace(valueD, " ", "")
            Else
                searchName = valueC
                searchSurname = valueD
            End If

            If ignoreCase Then
                searchName = LCase(searchName)
                searchSurname = LCase(searchSurname)
            End If

            lastRowArchiv = wsArchiv.Cells(wsArchiv.rows.Count, "C").End(xlUp).row

            foundRow = 0
            For i = 2 To lastRowArchiv
                compareC = wsArchiv.Cells(i, "C").value
                compareD = wsArchiv.Cells(i, "D").value

                If ignoreSpaces Then
                    compareC = Replace(compareC, " ", "")
                    compareD = Replace(compareD, " ", "")
                End If

                If ignoreCase Then
                    compareC = LCase(compareC)
                    compareD = LCase(compareD)
                End If

                If compareC = searchName And compareD = searchSurname Then
                    foundRow = i
                    Exit For
                End If
            Next i

            If foundRow > 0 Then
                wsKinder.Cells(currentRow, "B").value = wsArchiv.Cells(foundRow, "B").value
                wsKinder.Cells(currentRow, "G").value = wsArchiv.Cells(foundRow, "G").value
                wsKinder.Cells(currentRow, "H").value = wsArchiv.Cells(foundRow, "H").value
                wsKinder.Cells(currentRow, "K").value = wsArchiv.Cells(foundRow, "K").value
                wsKinder.Cells(currentRow, "L").value = wsArchiv.Cells(foundRow, "L").value
                wsKinder.Cells(currentRow, "O").value = wsArchiv.Cells(foundRow, "O").value
                wsKinder.Cells(currentRow, "S").value = wsArchiv.Cells(foundRow, "S").value
                wsKinder.Cells(currentRow, "T").value = wsArchiv.Cells(foundRow, "T").value
            End If
        End If
    End If
End Sub

Private Function IsValidFormat(ByVal textVal As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    With regex
        .pattern = "^\d\.\s\d{4}$"
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
