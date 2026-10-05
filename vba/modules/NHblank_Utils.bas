Attribute VB_Name = "NHblank_Utils"
Option Explicit

Private Const PRIMARY_TEMPLATE_NAME As String = "Shablon.docx"
Private Const LEGACY_TEMPLATE_NAME As String = "Shablon.xlsx"

' Get German month name without umlauts
Public Function GetGermanMonth(ByVal monthNumber As Integer) As String
    Select Case monthNumber
        Case 1: GetGermanMonth = "Januar"
        Case 2: GetGermanMonth = "Februar"
        Case 3: GetGermanMonth = "Maerz"
        Case 4: GetGermanMonth = "April"
        Case 5: GetGermanMonth = "Mai"
        Case 6: GetGermanMonth = "Juni"
        Case 7: GetGermanMonth = "Juli"
        Case 8: GetGermanMonth = "August"
        Case 9: GetGermanMonth = "September"
        Case 10: GetGermanMonth = "Oktober"
        Case 11: GetGermanMonth = "November"
        Case 12: GetGermanMonth = "Dezember"
        Case Else: GetGermanMonth = ""
    End Select
End Function

' Format date range as "von DD.MM.YYYY bis DD.MM.YYYY"
Public Function FormatDateRange(ByVal dateFrom As Date, ByVal dateTo As Date) As String
    FormatDateRange = "Bewilligungszeitraum von " & Format(dateFrom, "DD.MM.YYYY") & " bis " & Format(dateTo, "DD.MM.YYYY")
End Function

' Create safe file name from last name, first name and discipline
' Replaces spaces with underscores and removes invalid characters
Public Function CreateFileName( _
    ByVal lastName As String, _
    ByVal firstName As String, _
    ByVal discipline As String _
) As String
    
    Dim fileName As String
    
    ' Build file name
    fileName = lastName & "_" & firstName & "_" & discipline
    
    ' Replace spaces with underscores
    fileName = Replace(fileName, " ", "_")
    
    ' Remove invalid file name characters
    fileName = RemoveInvalidChars(fileName)
    
    CreateFileName = fileName
End Function

' Create a monthly file name and optionally add a marker such as SA.
' Example:
'   Mustermann_Max_Mathe_Juli_2026
'   Mustermann_Max_Mathe_SA_Juli_2026
Public Function CreateFileNameWithPeriod( _
    ByVal lastName As String, _
    ByVal firstName As String, _
    ByVal discipline As String, _
    ByVal referenceDate As Date, _
    ByVal marker As String _
) As String

    Dim fileName As String

    fileName = lastName & "_" & firstName & "_" & discipline

    If Len(Trim$(marker)) > 0 Then
        fileName = fileName & "_" & Trim$(marker)
    End If

    fileName = fileName & "_" & _
               GetGermanMonth(Month(referenceDate)) & "_" & _
               CStr(Year(referenceDate))

    fileName = Replace(fileName, " ", "_")
    fileName = RemoveInvalidChars(fileName)

    CreateFileNameWithPeriod = fileName
End Function

' Create safe folder name from teacher name
' Replaces spaces with underscores and removes invalid characters
Public Function CreateSafeFolderName(ByVal teacherName As String) As String
    Dim folderName As String
    
    ' Replace spaces with underscores
    folderName = Replace(teacherName, " ", "_")
    
    ' Remove invalid folder name characters
    folderName = RemoveInvalidChars(folderName)
    
    CreateSafeFolderName = folderName
End Function

' Remove characters that are invalid in Windows file names
Private Function RemoveInvalidChars(ByVal fileName As String) As String
    Dim invalidChars As Variant
    Dim i As Long
    Dim result As String
    
    ' List of invalid characters in Windows file names
    invalidChars = Array("/", "\", ":", "*", "?", """", "<", ">", "|")
    
    result = fileName
    
    ' Remove each invalid character
    For i = LBound(invalidChars) To UBound(invalidChars)
        result = Replace(result, invalidChars(i), "")
    Next i
    
    RemoveInvalidChars = result
End Function

' Debug function to show path information.
Public Sub ShowPathDebugInfo()
    Dim msg As String

    msg = "Debug Info:" & vbCrLf & vbCrLf
    msg = msg & "ThisWorkbook.Path: " & ThisWorkbook.Path & vbCrLf
    msg = msg & "ThisWorkbook.FullName: " & ThisWorkbook.fullName & vbCrLf
    msg = msg & "OneDriveCommercial: " & _
                Environ$("OneDriveCommercial") & vbCrLf
    msg = msg & "OneDriveConsumer: " & _
                Environ$("OneDriveConsumer") & vbCrLf
    msg = msg & "OneDrive: " & Environ$("OneDrive") & vbCrLf
    msg = msg & vbCrLf & "GetLocalPath result: " & _
                GetLocalPath(ThisWorkbook.Path)

    MsgBox msg, vbInformation, "Path Debug Info"
End Sub

' Return the local folder which contains the running workbook.
'
' Excel may expose a OneDrive/SharePoint workbook as an HTTPS URL even when
' the files are synchronized locally. For such URLs, rebuild the relative
' folder below Documents and combine it with registered OneDrive sync roots.
' If no trustworthy mapping exists, return an empty string so the existing
' manual file pickers remain the fallback.
Public Function GetLocalPath(ByVal workbookPath As String) As String
    Dim fso As Object
    Dim localPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    workbookPath = Trim$(workbookPath)

    If Len(workbookPath) = 0 Then Exit Function

    ' Prefer FullName when Office exposes a genuine local file path there.
    localPath = ExtractLocalPathFromFullName(ThisWorkbook.fullName)
    If Len(localPath) > 0 Then
        If fso.FolderExists(localPath) Then
            GetLocalPath = RemoveTrailingSeparator(localPath)
            Exit Function
        End If
    End If

    If Not IsHttpUrl(workbookPath) Then
        If fso.FolderExists(workbookPath) Then
            GetLocalPath = RemoveTrailingSeparator(workbookPath)
        End If
        Exit Function
    End If

    localPath = ResolveOneDriveUrlToLocalFolder(workbookPath)
    If Len(localPath) > 0 Then
        GetLocalPath = localPath
    End If
End Function

Private Function IsHttpUrl(ByVal value As String) As Boolean
    IsHttpUrl = _
        (LCase$(Left$(value, 7)) = "http://") Or _
        (LCase$(Left$(value, 8)) = "https://")
End Function

' Extract the directory from a local FullName; cloud URLs return empty.
Private Function ExtractLocalPathFromFullName( _
    ByVal fullName As String _
) As String

    Dim lastBackslash As Long

    If IsHttpUrl(fullName) Then Exit Function

    If InStr(1, fullName, ":\", vbTextCompare) = 0 And _
       Left$(fullName, 2) <> "\\" Then
        Exit Function
    End If

    lastBackslash = InStrRev(fullName, "\")
    If lastBackslash > 0 Then
        ExtractLocalPathFromFullName = _
            Left$(fullName, lastBackslash - 1)
    End If
End Function

Private Function ResolveOneDriveUrlToLocalFolder( _
    ByVal oneDriveUrl As String _
) As String

    Dim roots As Collection
    Dim relativePaths As Collection
    Dim fso As Object
    Dim rootPath As Variant
    Dim relativePath As Variant
    Dim candidate As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set roots = GetOneDriveSyncRoots(fso)
    Set relativePaths = GetUrlRelativePathCandidates(oneDriveUrl)

    For Each relativePath In relativePaths
        For Each rootPath In roots
            candidate = CombineRootAndRelative( _
                CStr(rootPath), CStr(relativePath), fso)

            If IsExpectedWorkbookFolder(candidate, fso) Then
                ResolveOneDriveUrlToLocalFolder = candidate
                Exit Function
            End If
        Next rootPath
    Next relativePath
End Function

Private Function GetOneDriveSyncRoots(ByVal fso As Object) As Collection
    Dim result As Collection
    Dim seen As Object
    Dim accountName As String
    Dim registryPath As String
    Dim registryValue As String
    Dim accountIndex As Long
    Dim userProfile As String
    Dim profileFolder As Object
    Dim subFolder As Object

    Set result = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    ' Environment variables are fast and normally identify the active roots.
    AddSyncRoot result, seen, Environ$("OneDriveCommercial"), fso
    AddSyncRoot result, seen, Environ$("OneDriveConsumer"), fso
    AddSyncRoot result, seen, Environ$("OneDrive"), fso

    ' Registry values cover additional business accounts.
    accountName = "Personal"
    registryPath = "HKEY_CURRENT_USER\Software\Microsoft\OneDrive\" & _
                   "Accounts\" & accountName & "\UserFolder"
    registryValue = TryReadRegistryString(registryPath)
    AddSyncRoot result, seen, registryValue, fso

    For accountIndex = 1 To 9
        accountName = "Business" & CStr(accountIndex)
        registryPath = "HKEY_CURRENT_USER\Software\Microsoft\OneDrive\" & _
                       "Accounts\" & accountName & "\UserFolder"
        registryValue = TryReadRegistryString(registryPath)
        AddSyncRoot result, seen, registryValue, fso
    Next accountIndex

    ' Last local-only fallback: OneDrive folders directly below USERPROFILE.
    userProfile = Environ$("USERPROFILE")
    If fso.FolderExists(userProfile) Then
        Set profileFolder = fso.GetFolder(userProfile)
        For Each subFolder In profileFolder.SubFolders
            If LCase$(Left$(subFolder.name, 8)) = "onedrive" Then
                AddSyncRoot result, seen, subFolder.Path, fso
            End If
        Next subFolder
    End If

    Set GetOneDriveSyncRoots = result
End Function

Private Sub AddSyncRoot( _
    ByVal result As Collection, _
    ByVal seen As Object, _
    ByVal candidate As String, _
    ByVal fso As Object _
)

    Dim normalized As String

    normalized = RemoveTrailingSeparator(Trim$(candidate))
    If Len(normalized) = 0 Then Exit Sub
    If Not fso.FolderExists(normalized) Then Exit Sub

    If Not seen.exists(normalized) Then
        seen.Add normalized, True
        result.Add normalized
    End If
End Sub

Private Function TryReadRegistryString( _
    ByVal registryPath As String _
) As String

    Dim shell As Object
    Dim value As Variant

    On Error Resume Next
    Set shell = CreateObject("WScript.Shell")
    value = shell.RegRead(registryPath)
    If Err.Number = 0 Then
        TryReadRegistryString = CStr(value)
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetUrlRelativePathCandidates( _
    ByVal oneDriveUrl As String _
) As Collection

    Dim result As Collection
    Dim seen As Object
    Dim decodedUrl As String
    Dim urlPath As String
    Dim markers As Variant
    Dim marker As Variant
    Dim relativePath As String
    Dim trimmedPath As String
    Dim parts As Variant
    Dim takeCount As Long
    Dim maxTakeCount As Long
    Dim firstPart As Long
    Dim i As Long

    Set result = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    decodedUrl = UrlDecodeUtf8(StripUrlQueryAndFragment(oneDriveUrl))
    urlPath = GetUrlPath(decodedUrl)

    markers = Array( _
        "/Documents/", _
        "/Shared Documents/", _
        "/Freigegebene Dokumente/", _
        "/Dokumente/")

    For Each marker In markers
        relativePath = PathAfterMarker(urlPath, CStr(marker))
        AddRelativeCandidate result, seen, relativePath
    Next marker

    ' OneDrive Personal URLs use /<cid>/<relative path>.
    If InStr(1, decodedUrl, "d.docs.live.net", vbTextCompare) > 0 Then
        trimmedPath = TrimSlashes(urlPath)
        parts = Split(trimmedPath, "/")
        If UBound(parts) >= 1 Then
            relativePath = JoinParts(parts, 1, UBound(parts))
            AddRelativeCandidate result, seen, relativePath
        End If
    End If

    ' Conservative suffix fallback for unusual SharePoint URL namespaces.
    trimmedPath = TrimSlashes(urlPath)
    If Len(trimmedPath) > 0 Then
        parts = Split(trimmedPath, "/")
        maxTakeCount = UBound(parts) + 1
        If maxTakeCount > 6 Then maxTakeCount = 6

        For takeCount = maxTakeCount To 1 Step -1
            firstPart = UBound(parts) - takeCount + 1
            relativePath = JoinParts(parts, firstPart, UBound(parts))
            AddRelativeCandidate result, seen, relativePath
        Next takeCount
    End If

    Set GetUrlRelativePathCandidates = result
End Function

Private Sub AddRelativeCandidate( _
    ByVal result As Collection, _
    ByVal seen As Object, _
    ByVal relativePath As String _
)

    Dim normalized As String

    normalized = TrimSlashes(Replace(relativePath, "\", "/"))
    If Len(normalized) = 0 Then Exit Sub

    If Not seen.exists(normalized) Then
        seen.Add normalized, True
        result.Add normalized
    End If
End Sub

Private Function PathAfterMarker( _
    ByVal urlPath As String, _
    ByVal marker As String _
) As String

    Dim markerPosition As Long

    markerPosition = InStr(1, urlPath, marker, vbTextCompare)
    If markerPosition > 0 Then
        PathAfterMarker = Mid$( _
            urlPath, markerPosition + Len(marker))
    End If
End Function

Private Function GetUrlPath(ByVal url As String) As String
    Dim schemePosition As Long
    Dim pathPosition As Long

    schemePosition = InStr(1, url, "://", vbTextCompare)
    If schemePosition = 0 Then Exit Function

    pathPosition = InStr(schemePosition + 3, url, "/")
    If pathPosition > 0 Then
        GetUrlPath = Mid$(url, pathPosition)
    End If
End Function

Private Function StripUrlQueryAndFragment(ByVal url As String) As String
    Dim cutPosition As Long
    Dim queryPosition As Long
    Dim fragmentPosition As Long

    cutPosition = Len(url) + 1
    queryPosition = InStr(1, url, "?", vbBinaryCompare)
    fragmentPosition = InStr(1, url, "#", vbBinaryCompare)

    If queryPosition > 0 And queryPosition < cutPosition Then
        cutPosition = queryPosition
    End If
    If fragmentPosition > 0 And fragmentPosition < cutPosition Then
        cutPosition = fragmentPosition
    End If

    StripUrlQueryAndFragment = Left$(url, cutPosition - 1)
End Function

Private Function TrimSlashes(ByVal value As String) As String
    Do While Left$(value, 1) = "/" Or Left$(value, 1) = "\"
        value = Mid$(value, 2)
    Loop

    Do While Right$(value, 1) = "/" Or Right$(value, 1) = "\"
        value = Left$(value, Len(value) - 1)
    Loop

    TrimSlashes = value
End Function

Private Function JoinParts( _
    ByVal parts As Variant, _
    ByVal firstIndex As Long, _
    ByVal lastIndex As Long _
) As String

    Dim i As Long
    Dim result As String

    For i = firstIndex To lastIndex
        If Len(result) > 0 Then result = result & "\"
        result = result & CStr(parts(i))
    Next i

    JoinParts = result
End Function

Private Function CombineRootAndRelative( _
    ByVal rootPath As String, _
    ByVal relativePath As String, _
    ByVal fso As Object _
) As String

    relativePath = Replace(relativePath, "/", "\")
    relativePath = TrimSlashes(relativePath)

    If Len(relativePath) = 0 Then
        CombineRootAndRelative = RemoveTrailingSeparator(rootPath)
    Else
        CombineRootAndRelative = fso.BuildPath( _
            RemoveTrailingSeparator(rootPath), relativePath)
    End If
End Function

Private Function IsExpectedWorkbookFolder( _
    ByVal folderPath As String, _
    ByVal fso As Object _
) As Boolean

    Dim workbookName As String

    If Len(folderPath) = 0 Then Exit Function
    If Not fso.FolderExists(folderPath) Then Exit Function

    On Error Resume Next
    workbookName = ThisWorkbook.name
    On Error GoTo 0

    If Len(workbookName) > 0 Then
        If fso.FileExists(fso.BuildPath(folderPath, workbookName)) Then
            IsExpectedWorkbookFolder = True
            Exit Function
        End If
    End If

    IsExpectedWorkbookFolder = _
        fso.FileExists(fso.BuildPath( _
            folderPath, PRIMARY_TEMPLATE_NAME)) Or _
        fso.FileExists(fso.BuildPath( _
            folderPath, LEGACY_TEMPLATE_NAME))
End Function

Private Function RemoveTrailingSeparator(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)

    Do While Len(folderPath) > 3 And _
             (Right$(folderPath, 1) = "\" Or _
              Right$(folderPath, 1) = "/")
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    Loop

    RemoveTrailingSeparator = folderPath
End Function

' Decode percent-encoded UTF-8 URL paths without an Office reference.
Private Function UrlDecodeUtf8(ByVal encodedText As String) As String
    Dim bytes() As Byte
    Dim byteCount As Long
    Dim i As Long
    Dim currentChar As String
    Dim codePoint As Long
    Dim stream As Object

    On Error GoTo SimpleFallback

    ReDim bytes(0 To Len(encodedText) * 3)
    i = 1

    Do While i <= Len(encodedText)
        currentChar = Mid$(encodedText, i, 1)

        If currentChar = "%" And i + 2 <= Len(encodedText) And _
           IsHexPair(Mid$(encodedText, i + 1, 2)) Then
            bytes(byteCount) = CByte( _
                CLng("&H" & Mid$(encodedText, i + 1, 2)))
            byteCount = byteCount + 1
            i = i + 3
        Else
            codePoint = AscW(currentChar)
            If codePoint < 0 Then codePoint = codePoint + 65536
            AppendUtf8CodePoint bytes, byteCount, codePoint
            i = i + 1
        End If
    Loop

    If byteCount = 0 Then Exit Function
    ReDim Preserve bytes(0 To byteCount - 1)

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "utf-8"
    UrlDecodeUtf8 = stream.ReadText
    stream.Close
    Exit Function

SimpleFallback:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    UrlDecodeUtf8 = Replace(encodedText, "%20", " ", _
                            Compare:=vbTextCompare)
End Function

Private Function IsHexPair(ByVal value As String) As Boolean
    Dim parsed As Long

    If Len(value) <> 2 Then Exit Function

    On Error GoTo NotHex
    parsed = CLng("&H" & value)
    IsHexPair = (parsed >= 0 And parsed <= 255)
    Exit Function

NotHex:
    IsHexPair = False
End Function

Private Sub AppendUtf8CodePoint( _
    ByRef bytes() As Byte, _
    ByRef byteCount As Long, _
    ByVal codePoint As Long _
)

    If codePoint < 128 Then
        bytes(byteCount) = CByte(codePoint)
        byteCount = byteCount + 1
    ElseIf codePoint < 2048 Then
        bytes(byteCount) = CByte(192 Or (codePoint \ 64))
        bytes(byteCount + 1) = CByte(128 Or (codePoint And 63))
        byteCount = byteCount + 2
    Else
        bytes(byteCount) = CByte(224 Or (codePoint \ 4096))
        bytes(byteCount + 1) = CByte( _
            128 Or ((codePoint \ 64) And 63))
        bytes(byteCount + 2) = CByte(128 Or (codePoint And 63))
        byteCount = byteCount + 3
    End If
End Sub

