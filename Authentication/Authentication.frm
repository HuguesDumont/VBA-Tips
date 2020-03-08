VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Authentication 
   Caption         =   "Connect"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6780
   OleObjectBlob   =   "Authentication.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "authentication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Pre-made window for simple connection using login and password"
Option Explicit

Private Sub cancel_Click()
Attribute cancel_Click.VB_Description = "Empty login and password fields"
    Me.login.value = ""
    Me.pass.value = ""
End Sub

Private Sub UserForm_Activate()
Attribute UserForm_Activate.VB_Description = "Show authentification window. Auto-updating and displaying date and time"
    While True
        Me.dateTime.Caption = CStr(Now)
        DoEvents
    Wend
End Sub

Private Sub validate_Click()
    
End Sub
