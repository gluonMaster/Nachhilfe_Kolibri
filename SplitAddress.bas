Attribute VB_Name = "SplitAddress"
Option Explicit

Public Sub SplitAddressIntoStreetAndCity()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim addressValue As String
    Dim postalCode As String
    Dim leftPart As String
    Dim rightPart As String
    Dim codePosition As Long
    Dim regex As Object
    Dim match As Object
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo Cleanup
    
    ' Set the worksheet (adjust if needed)
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Find the last used row in column S
    If Application.WorksheetFunction.CountA(ws.Columns("S")) < 5 Then
        ' No data starting from row 5
        GoTo Cleanup
    Else
        lastRow = ws.Cells(ws.Rows.Count, "S").End(xlUp).row
    End If
    
    ' Create a RegExp object to find a 5-digit postal code
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\b\d{5}\b"
    regex.Global = False
    regex.ignoreCase = True
    
    For currentRow = 5 To lastRow
        addressValue = Trim(CStr(ws.Cells(currentRow, "S").value))
        If addressValue <> "" Then
            ' Try to find the 5-digit postal code
            If regex.Test(addressValue) Then
                Set match = regex.Execute(addressValue)(0)
                
                ' Extract the postal code and split the address
                postalCode = match.value
                ' match.FirstIndex is zero-based, Mid function is 1-based
                codePosition = match.FirstIndex + 1
                
                leftPart = Left(addressValue, codePosition - 1)
                ' Add the length of postalCode and start from next character
                rightPart = mid(addressValue, codePosition + Len(postalCode))
                
                ' Clean up punctuation and spaces around the postal code
                leftPart = CleanTrailingPunctuation(leftPart)
                rightPart = CleanLeadingPunctuation(rightPart)
                
                ' If after cleaning rightPart does not start with a space and is not empty, add one
                If Len(rightPart) > 0 And Left(rightPart, 1) <> " " Then
                    rightPart = " " & rightPart
                End If
                
                ' Put the cleaned parts back into cells
                ws.Cells(currentRow, "S").value = Trim(leftPart)
                ws.Cells(currentRow, "T").value = postalCode & rightPart
            Else
                ' If no 5-digit code found, do nothing or handle differently if needed
            End If
        End If
    Next currentRow
    
Cleanup:
    ' Restore settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Private Function CleanTrailingPunctuation(ByVal textVal As String) As String
    ' Remove trailing spaces, commas, semicolons, periods
    Dim temp As String
    temp = RTrim(textVal)
    Do While Len(temp) > 0 And (Right(temp, 1) = "," Or Right(temp, 1) = ";" Or Right(temp, 1) = "." Or Right(temp, 1) = " ")
        temp = Left(temp, Len(temp) - 1)
        temp = RTrim(temp)
    Loop
    CleanTrailingPunctuation = RTrim(temp)
End Function

Private Function CleanLeadingPunctuation(ByVal textVal As String) As String
    ' Remove leading spaces, commas, semicolons, periods
    Dim temp As String
    temp = LTrim(textVal)
    Do While Len(temp) > 0 And (Left(temp, 1) = "," Or Left(temp, 1) = ";" Or Left(temp, 1) = "." Or Left(temp, 1) = " ")
        temp = mid(temp, 2)
        temp = LTrim(temp)
    Loop
    CleanLeadingPunctuation = LTrim(temp)
End Function

