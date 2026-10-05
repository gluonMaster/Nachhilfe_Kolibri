VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NHblank_AppState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' NHblank_AppState
' Safely captures, optimizes, and restores Excel application state.
' Guarantees state restoration even if an error occurs during macro execution.
' =============================================================================

' Stored application state
Private m_ScreenUpdating As Boolean
Private m_Calculation As XlCalculation
Private m_EnableEvents As Boolean
Private m_DisplayAlerts As Boolean
Private m_StatusBar As Variant

' Flag to track if state was captured (prevents double-restore issues)
Private m_IsCaptured As Boolean

' -----------------------------------------------------------------------------
' Capture: Saves the current Excel application state
' -----------------------------------------------------------------------------
Public Sub Capture()
    If m_IsCaptured Then Exit Sub
    
    With Application
        m_ScreenUpdating = .ScreenUpdating
        m_Calculation = .Calculation
        m_EnableEvents = .EnableEvents
        m_DisplayAlerts = .DisplayAlerts
        m_StatusBar = .StatusBar
    End With
    
    m_IsCaptured = True
End Sub

' -----------------------------------------------------------------------------
' OptimizeForRun: Sets Excel to optimal state for macro performance
' Call Capture() before this to save original state
' -----------------------------------------------------------------------------
Public Sub OptimizeForRun()
    ' Auto-capture if not done yet
    If Not m_IsCaptured Then Capture
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With
End Sub

' -----------------------------------------------------------------------------
' Restore: Restores Excel application state to captured values
' Idempotent: safe to call multiple times
' -----------------------------------------------------------------------------
Public Sub Restore()
    If Not m_IsCaptured Then Exit Sub
    
    With Application
        .ScreenUpdating = m_ScreenUpdating
        .Calculation = m_Calculation
        .EnableEvents = m_EnableEvents
        .DisplayAlerts = m_DisplayAlerts
        
        ' Restore StatusBar: False clears custom message, string restores it
        If VarType(m_StatusBar) = vbBoolean Then
            .StatusBar = False
        ElseIf VarType(m_StatusBar) = vbString Then
            .StatusBar = m_StatusBar
        Else
            .StatusBar = False
        End If
    End With
    
    m_IsCaptured = False
End Sub

' -----------------------------------------------------------------------------
' Class_Terminate: Automatic cleanup when object is destroyed
' Ensures state is restored even if Restore() was not explicitly called
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    Restore
End Sub
