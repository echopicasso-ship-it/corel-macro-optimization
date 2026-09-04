VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbort
   Caption         =   "Processing..."
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "frmAbort.frx":0000
   StartUpPosition =   1
   WhatsThisHelp   =   0
End
Attribute VB_Name = "frmAbort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdAbort_Click()
    ' Устанавливаем флаг отмены
    GlobalMacros.CancelRequested = True
    
    ' Разгружаем форму
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.lblStatus.Caption = "Processing..."
    Me.cmdAbort.Caption = "ABORT"
    Me.cmdAbort.Enabled = True
End Sub
