Attribute VB_Name = "AdressImport"
Option Explicit

Public Sub LoadAddressesFromExternalFile()
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim filePath As String
    Dim lastRowSource As Long, lastRowDest As Long
    Dim currentRow As Long
    Dim dictAddresses As Object
    Dim keyValue As String
    Dim cell As Range
    Dim countDict As Object
    Dim i As Long
    
    ' Ask user to select the external file
    filePath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , "Select the external file")
    If filePath = "False" Then Exit Sub ' User canceled
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    
    ' Open the external file in read-only mode
    Set wbSource = Workbooks.Open(filePath, ReadOnly:=True)
    
    On Error GoTo Cleanup
    ' Set the source worksheet "Kartei"
    Set wsSource = wbSource.Sheets("Kartei")
    
    ' Create a dictionary for addresses
    Set dictAddresses = CreateObject("Scripting.Dictionary")
    ' This dictionary will store the second occurrence address for each number.
    ' Key: Number (string), Value: Address from column F
    
    ' Create a dictionary to count occurrences
    Set countDict = CreateObject("Scripting.Dictionary")
    
    ' Find the last row in source worksheet
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).row
    
    ' Loop through source data
    ' Column A: Numbers in "X. XXXX" format
    ' Column F: Addresses
    For i = 2 To lastRowSource ' Assuming header in row 1
        keyValue = Trim(CStr(wsSource.Cells(i, "A").value))
        
        ' Increase occurrence count for this key
        If Not countDict.exists(keyValue) Then
            countDict(keyValue) = 1
        Else
            countDict(keyValue) = countDict(keyValue) + 1
        End If
        
        ' If this is the second occurrence of the key, store its address from column F
        If countDict(keyValue) = 2 Then
            ' Only store if it is the second occurrence
            dictAddresses(keyValue) = wsSource.Cells(i, "F").value
        End If
    Next i
    
    ' Go back to the destination worksheet (assuming this code is in the same workbook)
    Set wsDest = ThisWorkbook.ActiveSheet
    
    ' Find the last used row in column B in the destination sheet
    If Application.WorksheetFunction.CountA(wsDest.Columns("B")) < 5 Then
        ' No data starting from row 5
        GoTo Cleanup
    Else
        lastRowDest = wsDest.Cells(wsDest.Rows.Count, "B").End(xlUp).row
    End If
    
    ' Loop through destination rows starting from 5
    For currentRow = 5 To lastRowDest
        keyValue = Trim(CStr(wsDest.Cells(currentRow, "B").value))
        ' If we have the address for this key (second occurrence from source)
        If dictAddresses.exists(keyValue) Then
            wsDest.Cells(currentRow, "S").value = dictAddresses(keyValue)
        Else
            ' If needed, handle the case when address not found
            ' e.g., highlight cell, show message, etc.
        End If
    Next currentRow
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    
Cleanup:
    If Not wbSource Is Nothing Then
        wbSource.Close SaveChanges:=False
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

