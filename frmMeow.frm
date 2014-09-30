VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMeow 
   Caption         =   "Anti-Meowdler"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3105
   OleObjectBlob   =   "frmMeow.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMeow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnMeow_Click()
    Meowdle.BumpIntervalMinutes = CInt(Me.txtThreshold)
    Meowdle.QueueBump
    Meowdle.QueueStatusUpdate
    If Me.chkCatNap.Value Then
        Meowdle.NapPeriodHours = CInt(Me.txtNapThreshold)
        Meowdle.QueueNap
        Me.txtNapThreshold.Enabled = False
    End If
    Me.btnMAOU.Enabled = True
    Me.btnMeow.Enabled = False
    Me.chkCatNap.Enabled = False
    Me.txtThreshold.Enabled = False
End Sub
Private Sub btnMAOU_Click()
    If Me.chkCatNap.Value Then
        Meowdle.CancelNap
        Me.txtNapThreshold.Enabled = True
    End If
    Meowdle.Nap
End Sub
Private Sub btnMeownimize_Click()
    Meowdle.Miew
End Sub
Private Sub chkCatNap_Click()
    If Me.chkCatNap.Value Then
        Me.txtNapThreshold.Enabled = True
    Else
        Me.txtNapThreshold.Enabled = False
    End If
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Meowdle.Miew True
End Sub

