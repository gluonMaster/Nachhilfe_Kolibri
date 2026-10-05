VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SourceRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ====================================================================
' Class: SourceRecord
' Description: Holds source record data
' ====================================================================

Public RowNumber As Long
Public ID As String
Public lastName As String
Public firstName As String
Public fullNameCleaned As String
Public subject As String
Public value As Variant
