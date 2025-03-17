Attribute VB_Name = "LoadTableClear"
Option Explicit

Sub ClearEducationalLoadTable()
    On Error GoTo ErrorHandler
    
    ' Declare necessary variables
    Dim targetWs As Worksheet
    Dim lastRowAtoAN As Long
    Dim lastRowAUtoAV As Long
    Dim lastRow As Long
    Dim clearRange1 As Range
    Dim clearRange2 As Range
    Dim backgroundRange As Range
    
    ' **1. Set the target worksheet**
    Set targetWs = ActiveSheet ' Change to a specific sheet if necessary, e.g., ThisWorkbook.Sheets("SheetName")
    
    ' **2. Determine the last used row in the specified columns**
    ' Find the last used row in columns A to AN
    lastRowAtoAN = targetWs.Cells(targetWs.Rows.Count, "A").End(xlUp).row
    lastRowAtoAN = Application.WorksheetFunction.Max(lastRowAtoAN, targetWs.Cells(targetWs.Rows.Count, "AN").End(xlUp).row)
    
    ' Find the last used row in columns AU to AV
    lastRowAUtoAV = targetWs.Cells(targetWs.Rows.Count, "AU").End(xlUp).row
    lastRowAUtoAV = Application.WorksheetFunction.Max(lastRowAUtoAV, targetWs.Cells(targetWs.Rows.Count, "AV").End(xlUp).row)
    
    ' Determine the overall last row to clear
    lastRow = Application.WorksheetFunction.Max(lastRowAtoAN, lastRowAUtoAV)
    
    ' **3. Define the ranges to clear**
    If lastRow >= 11 Then
        Set clearRange1 = targetWs.Range("A11:AN" & lastRow)
        Set clearRange2 = targetWs.Range("AU11:AY" & lastRow)
        Set backgroundRange = targetWs.Range("I11:AN" & lastRow)
    Else
        ' If there are no data rows beyond row 10, exit the macro
        MsgBox "There are no records to clear.", vbInformation
        Exit Sub
    End If
    
    ' **4. Improve performance**
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' **5. Clear the specified ranges**
    clearRange1.ClearContents
    clearRange2.ClearContents
    backgroundRange.Interior.ColorIndex = xlNone
    
    ' **6. Restore settings**
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' **7. Inform the user of successful completion**
    MsgBox "Educational Load table has been cleared successfully.", vbInformation
    Exit Sub

ErrorHandler:
    ' **8. Handle any unexpected errors**
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

