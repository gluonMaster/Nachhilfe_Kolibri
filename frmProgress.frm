VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "Progress"
   ClientHeight    =   2376
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9732
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cancelRequested As Boolean

Private Sub btnCancel_Click()
    cancelRequested = True
    btnCancel.Enabled = False ' Prevent multiple clicks
    btnCancel.Caption = "Cancelling..."
End Sub

