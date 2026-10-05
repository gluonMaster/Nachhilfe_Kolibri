Attribute VB_Name = "ubertNH_Finder"
Option Explicit

' ====================================================================
' Module: ubertNH_Finder
' Description: Find matching records in target worksheet
' ====================================================================

' Find matching record in target sheet
Public Function FindTargetRecord(ByVal targetSheet As Worksheet, _
                                ByRef srcRecord As SourceRecord, _
                                ByVal semester As Integer) As FindResult
    Dim result As FindResult
    Set result = New FindResult
    Dim lastRow As Long
    Dim currentRow As Long
    Dim targetId As String
    Dim targetFullName As String
    Dim targetFullNameCleaned As String
    Dim targetSubject As String
    Dim subjectColumn As String
    Dim matchedRows As Collection
    Dim matchRow As Variant
    
    ' Initialize result
    result.found = False
    result.targetRow = 0
    result.notFoundReason = ""
    
    ' Determine subject column based on semester
    If semester = 1 Then
        subjectColumn = ubertNH_Config.TGT_COL_SUBJECT_SEMESTER1
    Else
        subjectColumn = ubertNH_Config.TGT_COL_SUBJECT_SEMESTER2
    End If
    
    ' Get last row (check both columns A and B)
    lastRow = ubertNH_Utils.GetLastRowTwoColumns(targetSheet, _
                                                  ubertNH_Config.TGT_COL_ID, _
                                                  ubertNH_Config.TGT_COL_ID_CHECK)
    
    ' Step 1: Find all rows with matching ID
    Set matchedRows = New Collection
    
    For currentRow = ubertNH_Config.TGT_FIRST_DATA_ROW To lastRow
        ' Check if row is not empty (A or B)
        If Not IsEmpty(targetSheet.Range(ubertNH_Config.TGT_COL_ID & currentRow).value) Or _
           Not IsEmpty(targetSheet.Range(ubertNH_Config.TGT_COL_ID_CHECK & currentRow).value) Then
            
            targetId = Trim(CStr(targetSheet.Range(ubertNH_Config.TGT_COL_ID & currentRow).value))
            
            ' Compare ID
            If targetId = srcRecord.ID Then
                matchedRows.Add currentRow
            End If
        End If
    Next currentRow
    
    ' Check if any ID matches found
    If matchedRows.Count = 0 Then
        result.notFoundReason = "ID nicht gefunden"
        Set FindTargetRecord = result
        Exit Function
    End If
    
    ' Step 2: Among matched IDs, find rows with matching name
    Dim nameMatchedRows As Collection
    Set nameMatchedRows = New Collection
    
    For Each matchRow In matchedRows
        currentRow = CLng(matchRow)
        targetFullName = Trim(CStr(targetSheet.Range(ubertNH_Config.TGT_COL_FULLNAME & currentRow).value))
        targetFullNameCleaned = ubertNH_Utils.CleanText(targetFullName)
        
        ' Compare cleaned names
        If targetFullNameCleaned = srcRecord.fullNameCleaned Then
            nameMatchedRows.Add currentRow
        End If
    Next matchRow
    
    ' Check if any name matches found
    If nameMatchedRows.Count = 0 Then
        result.notFoundReason = "Name nicht gefunden (ID stimmt, aber Name stimmt nicht ueberein)"
        Set FindTargetRecord = result
        Exit Function
    End If
    
    ' Step 3: Among name-matched rows, find row with matching subject
    Dim subjectsChecked As String
    subjectsChecked = ""
    
    For Each matchRow In nameMatchedRows
        currentRow = CLng(matchRow)
        targetSubject = Trim(CStr(targetSheet.Range(subjectColumn & currentRow).value))
        
        ' Log subjects being compared (for debugging)
        If subjectsChecked <> "" Then subjectsChecked = subjectsChecked & "; "
        subjectsChecked = subjectsChecked & "Zeile " & currentRow & ": '" & targetSubject & "'"
        
        ' Check if target subject contains source subject
        If ubertNH_Utils.SubjectMatches(targetSubject, srcRecord.subject) Then
            result.found = True
            result.targetRow = currentRow
            Set FindTargetRecord = result
            Exit Function
        End If
    Next matchRow
    
    ' If we got here, subject was not found
    result.notFoundReason = "Fach nicht gefunden (ID und Name stimmen, aber Fach stimmt nicht ueberein). " & _
                           "Gesucht: '" & srcRecord.subject & "', Gefunden: " & subjectsChecked
    Set FindTargetRecord = result
End Function

