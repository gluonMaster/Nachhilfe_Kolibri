Attribute VB_Name = "NHblank_BgNummer"
Option Explicit

' =============================================================================
' NHblank_BgNummer
' Strict normalization and classification of Jobcenter / Sozialamt identifiers.
'
' Jobcenter:
'   12345//1234567
'   1234//1234567
' A single slash is normalized to a double slash.
'
' Sozialamt:
'   10301015032314
'   1.072.1.01.0346.3
'
' Numeric Excel values are recovered only when the integer is still exact
' (maximum 15 digits). Date-converted and ambiguous values are rejected.
' =============================================================================

Private Const MIN_SA_DIGITS As Long = 10
Private Const MAX_TEXT_SA_DIGITS As Long = 20
Private Const MAX_EXACT_EXCEL_DIGITS As Long = 15

' Normalize, classify and write the safe text value back to the source cell.
' Returns True only when the identifier can be reconstructed unambiguously.
Public Function NHblank_NormalizeAndProtectBgCell( _
    ByVal sourceCell As Range, _
    ByRef normalizedNumber As String, _
    ByRef isJobcenter As Boolean, _
    ByRef errorMsg As String _
) As Boolean

    Dim currentValue As String

    If Not NHblank_TryNormalizeBgCell( _
        sourceCell, normalizedNumber, isJobcenter, errorMsg) Then
        NHblank_NormalizeAndProtectBgCell = False
        Exit Function
    End If

    On Error GoTo WriteError

    currentValue = Trim$(CStr(sourceCell.Value2))
    If currentValue <> normalizedNumber Or sourceCell.NumberFormat <> "@" Then
        sourceCell.NumberFormat = "@"
        sourceCell.Value2 = normalizedNumber
    End If

    NHblank_NormalizeAndProtectBgCell = True
    Exit Function

WriteError:
    errorMsg = "BG-Nummer wurde erkannt, konnte aber nicht als Text " & _
               "geschuetzt werden: " & Err.Description
    NHblank_NormalizeAndProtectBgCell = False
End Function

' Read and normalize a BG number without changing the cell.
Public Function NHblank_TryNormalizeBgCell( _
    ByVal sourceCell As Range, _
    ByRef normalizedNumber As String, _
    ByRef isJobcenter As Boolean, _
    ByRef errorMsg As String _
) As Boolean

    Dim rawValue As Variant
    Dim cellValue As Variant
    Dim formatInvariant As String
    Dim formatLocal As String

    normalizedNumber = ""
    isJobcenter = False
    errorMsg = ""

    If sourceCell Is Nothing Then
        errorMsg = "BG-Nummer: Quellzelle fehlt."
        Exit Function
    End If

    On Error GoTo ReadError

    rawValue = sourceCell.Value2
    cellValue = sourceCell.Value
    formatInvariant = CStr(sourceCell.NumberFormat)
    formatLocal = CStr(sourceCell.NumberFormatLocal)

    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
        errorMsg = "BG-Nummer ist leer oder enthaelt einen Excel-Fehler."
        Exit Function
    End If

    If VarType(rawValue) = vbString Then
        NHblank_TryNormalizeBgCell = NormalizeTextValue( _
            CStr(rawValue), normalizedNumber, isJobcenter, errorMsg)
        Exit Function
    End If

    If IsNumeric(rawValue) Then
        If VarType(cellValue) = vbDate Or _
           IsDateLikeNumberFormat(formatInvariant) Or _
           IsDateLikeNumberFormat(formatLocal) Then
            errorMsg = "BG-Nummer wurde von Excel als Datum gespeichert. " & _
                       "Der urspruengliche Wert kann nicht eindeutig " & _
                       "wiederhergestellt werden."
            Exit Function
        End If

        NHblank_TryNormalizeBgCell = NormalizeNumericValue( _
            rawValue, normalizedNumber, isJobcenter, errorMsg)
        Exit Function
    End If

    errorMsg = "BG-Nummer hat einen nicht unterstuetzten Excel-Datentyp."
    Exit Function

ReadError:
    errorMsg = "BG-Nummer konnte nicht gelesen werden: " & Err.Description
    NHblank_TryNormalizeBgCell = False
End Function

Private Function NormalizeTextValue( _
    ByVal rawText As String, _
    ByRef normalizedNumber As String, _
    ByRef isJobcenter As Boolean, _
    ByRef errorMsg As String _
) As Boolean

    Dim cleaned As String
    Dim slashPosition As Long
    Dim digitCount As Long

    cleaned = Trim$(rawText)
    cleaned = Replace(cleaned, ChrW(160), "")
    cleaned = Replace(cleaned, " ", "")
    cleaned = Replace(cleaned, vbTab, "")
    cleaned = Replace(cleaned, vbCr, "")
    cleaned = Replace(cleaned, vbLf, "")

    If Left$(cleaned, 1) = "'" Then
        cleaned = Mid$(cleaned, 2)
    End If

    If Len(cleaned) = 0 Then
        errorMsg = "BG-Nummer ist leer."
        Exit Function
    End If

    ' A scientific-notation string may already have lost significant digits.
    ' Only a numeric Excel value up to 15 digits is recovered automatically.
    If InStr(1, cleaned, "E+", vbTextCompare) > 0 Or _
       InStr(1, cleaned, "E-", vbTextCompare) > 0 Then
        errorMsg = "BG-Nummer liegt nur als Text in Exponentialschreibweise " & _
                   "vor und kann nicht sicher rekonstruiert werden: " & rawText
        Exit Function
    End If

    ' Jobcenter semantics: 4 or 5 digits, one or two slashes, 7 digits.
    If RegexMatches(cleaned, "^\d{4,5}/{1,2}\d{7}$") Then
        slashPosition = InStr(1, cleaned, "/", vbBinaryCompare)
        normalizedNumber = Left$(cleaned, slashPosition - 1) & "//" & _
                           Mid$(cleaned, InStrRev(cleaned, "/") + 1)
        isJobcenter = True
        NormalizeTextValue = True
        Exit Function
    End If

    ' Sozialamt: uninterrupted digit sequence.
    If RegexMatches(cleaned, "^\d+$") Then
        digitCount = Len(cleaned)
        If digitCount >= MIN_SA_DIGITS And _
           digitCount <= MAX_TEXT_SA_DIGITS Then
            normalizedNumber = cleaned
            isJobcenter = False
            NormalizeTextValue = True
            Exit Function
        End If
    End If

    ' Sozialamt: several dot-separated numeric groups. Requiring at least
    ' four groups prevents an ordinary German date from being accepted.
    If RegexMatches(cleaned, "^\d+(\.\d+){3,}$") Then
        digitCount = Len(Replace(cleaned, ".", ""))
        If digitCount >= MIN_SA_DIGITS And _
           digitCount <= MAX_TEXT_SA_DIGITS Then
            normalizedNumber = cleaned
            isJobcenter = False
            NormalizeTextValue = True
            Exit Function
        End If
    End If

    errorMsg = "BG-Nummer hat kein eindeutig gueltiges Format: " & rawText & _
               ". Erwartet werden 12345//1234567, 12345/1234567, " & _
               "eine 10- bis 20-stellige Sozialamt-Nummer oder eine " & _
               "Sozialamt-Nummer mit mindestens vier Punktgruppen."
    NormalizeTextValue = False
End Function

Private Function NormalizeNumericValue( _
    ByVal rawValue As Variant, _
    ByRef normalizedNumber As String, _
    ByRef isJobcenter As Boolean, _
    ByRef errorMsg As String _
) As Boolean

    Dim numericValue As Double
    Dim recovered As String

    On Error GoTo NumericError

    numericValue = CDbl(rawValue)

    If numericValue < 0 Or numericValue <> Fix(numericValue) Then
        errorMsg = "Numerische BG-Nummer ist negativ oder enthaelt " & _
                   "Nachkommastellen und kann nicht sicher rekonstruiert werden."
        Exit Function
    End If

    recovered = Format$(numericValue, "0")

    If Len(recovered) > MAX_EXACT_EXCEL_DIGITS Then
        errorMsg = "Numerische BG-Nummer hat mehr als 15 Stellen. Excel kann " & _
                   "bereits Ziffern gerundet haben; eine sichere " & _
                   "Rekonstruktion ist nicht moeglich."
        Exit Function
    End If

    If Len(recovered) < MIN_SA_DIGITS Then
        errorMsg = "Numerische BG-Nummer ist zu kurz. Excel hat den Wert " & _
                   "moeglicherweise bereits in ein Datum oder eine andere " & _
                   "Zahl umgewandelt; der urspruengliche Wert kann nicht " & _
                   "eindeutig rekonstruiert werden: " & recovered
        Exit Function
    End If

    normalizedNumber = recovered
    isJobcenter = False
    NormalizeNumericValue = True
    Exit Function

NumericError:
    errorMsg = "Numerische BG-Nummer konnte nicht rekonstruiert werden: " & _
               Err.Description
    NormalizeNumericValue = False
End Function

Private Function RegexMatches(ByVal textValue As String, _
                              ByVal pattern As String) As Boolean
    Dim regex As Object

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = pattern
    regex.Global = False
    regex.IgnoreCase = True

    RegexMatches = regex.Test(textValue)
End Function

Private Function IsDateLikeNumberFormat(ByVal numberFormat As String) As Boolean
    Dim fmt As String

    fmt = LCase$(Trim$(numberFormat))

    If fmt = "" Or fmt = "general" Or fmt = "standard" Or fmt = "@" Then
        Exit Function
    End If

    ' English/German localized date tokens seen in Excel.
    If InStr(fmt, "yy") > 0 Or _
       InStr(fmt, "jj") > 0 Or _
       InStr(fmt, "dd") > 0 Or _
       InStr(fmt, "tt") > 0 Then
        IsDateLikeNumberFormat = True
    End If
End Function

' Lightweight, non-mutating test entry point for deployment verification.
' Returns "OK" when the normalization contract works as expected.
Public Function NHblank_BgNormalizationSelfTest() As String
    Dim normalized As String
    Dim isJobcenter As Boolean
    Dim errorMsg As String

    If Not AssertTextCase("12345//1234567", "12345//1234567", True) Then
        NHblank_BgNormalizationSelfTest = "Failed: standard Jobcenter"
        Exit Function
    End If

    If Not AssertTextCase("12345/1234567", "12345//1234567", True) Then
        NHblank_BgNormalizationSelfTest = "Failed: single slash"
        Exit Function
    End If

    If Not AssertTextCase("1234//1234567", "1234//1234567", True) Then
        NHblank_BgNormalizationSelfTest = "Failed: four-digit prefix"
        Exit Function
    End If

    If Not AssertTextCase("10301015032314", "10301015032314", False) Then
        NHblank_BgNormalizationSelfTest = "Failed: numeric Sozialamt text"
        Exit Function
    End If

    If Not AssertTextCase("1.072.1.01.0346.3", _
                          "1.072.1.01.0346.3", False) Then
        NHblank_BgNormalizationSelfTest = "Failed: dotted Sozialamt"
        Exit Function
    End If

    If NormalizeTextValue("1.03E+13", normalized, _
                          isJobcenter, errorMsg) Then
        NHblank_BgNormalizationSelfTest = _
            "Failed: ambiguous exponential text accepted"
        Exit Function
    End If

    If Not NormalizeNumericValue(10301015032314#, normalized, _
                                 isJobcenter, errorMsg) Or _
       normalized <> "10301015032314" Or isJobcenter Then
        NHblank_BgNormalizationSelfTest = _
            "Failed: exact numeric Sozialamt recovery"
        Exit Function
    End If

    NHblank_BgNormalizationSelfTest = "OK"
End Function

Private Function AssertTextCase(ByVal inputValue As String, _
                                ByVal expectedValue As String, _
                                ByVal expectedJobcenter As Boolean) As Boolean
    Dim normalized As String
    Dim isJobcenter As Boolean
    Dim errorMsg As String

    AssertTextCase = NormalizeTextValue( _
                         inputValue, normalized, isJobcenter, errorMsg) And _
                     normalized = expectedValue And _
                     isJobcenter = expectedJobcenter
End Function
