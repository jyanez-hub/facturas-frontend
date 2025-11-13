VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "UserForm1"
   ClientHeight    =   2424
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   6768
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub InitUI()
    Me.Caption = "Renombrando comprobantes…  (ESC para cancelar)"
    Me.lblText.Caption = "Preparando…"
    Me.fraBar.Caption = vbNullString
    Me.lblBar.Caption = vbNullString
    Me.lblBar.Move 2, 2, 0, Me.fraBar.InsideHeight - 4
    Me.Repaint
End Sub

Public Sub UpdateUI(ByVal pct As Double, ByVal statusText As String)
    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1
    Me.lblText.Caption = statusText
    Me.lblBar.Width = pct * (Me.fraBar.InsideWidth - 4)
    Me.Repaint
End Sub


