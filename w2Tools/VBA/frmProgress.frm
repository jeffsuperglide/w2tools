VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "Progress Bar"
   ClientHeight    =   1335
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5808
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
    ' Set the width of the progress bar to 0.
    frmProgress.LabelProgress.Width = 0

    ' Call the main subroutine.
    'Call Main
End Sub
