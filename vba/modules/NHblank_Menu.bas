Attribute VB_Name = "NHblank_Menu"
Option Explicit

' =============================================================================
' NHblank_Menu
' Public entry points for Excel buttons and ribbon assignments.
' Thin wrappers that delegate to actual implementation modules.
' =============================================================================

' -----------------------------------------------------------------------------
' NHblank_GenerateBlanks
' Entry point for blank generation. Calls NHblank_Main.GenerateBlanks.
' -----------------------------------------------------------------------------
Public Sub NHblank_GenerateBlanks()
    NHblank_Main.GenerateBlanks
End Sub

' -----------------------------------------------------------------------------
' NHblank_RunSyncKinderToBlanks
' Entry point for Kinder -> Kinder_Blanks synchronization.
' Copies active records from Kinder to Kinder_Blanks.
' -----------------------------------------------------------------------------
Public Sub NHblank_RunSyncKinderToBlanks()
    NHblank_Sync_KinderToBlanks.NHblank_SyncKinderToBlanks
End Sub

' -----------------------------------------------------------------------------
' NHblank_RunTransferSelectionKinderToBlanks
' Entry point for transferring selected rows from Kinder to Kinder_Blanks.
' -----------------------------------------------------------------------------
Public Sub NHblank_RunTransferSelectionKinderToBlanks()
    NHblank_TransferSelection.NHblank_TransferSelectedKinderRowsToBlanks
End Sub

' -----------------------------------------------------------------------------
' NHblank_RunSyncBlanksToKinder
' Entry point for Kinder_Blanks -> Kinder synchronization.
' Syncs changes from Kinder_Blanks back to Kinder.
' -----------------------------------------------------------------------------
Public Sub NHblank_RunSyncBlanksToKinder()
    NHblank_Sync_BlanksToKinder.NHblank_SyncBlanksToKinder
End Sub

' -----------------------------------------------------------------------------
' NHblank_RunDeactivateBlanksByT2
' Entry point for deactivating records on Kinder_Blanks based on T2 date.
' Marks records outside the reference month as inactive (gray font).
' -----------------------------------------------------------------------------
Public Sub NHblank_RunDeactivateBlanksByT2()
    NHblank_BlanksArchive.NHblank_DeactivateBlanksByT2
End Sub

' -----------------------------------------------------------------------------
' NHblank_RunFullAdminSync
' Entry point for full Admin -> Kinder -> Kinder_Blanks synchronization.
' -----------------------------------------------------------------------------
Public Sub NHblank_RunFullAdminSync()
    NHblank_AdminSync.NHblank_AdminSync_Run
End Sub
