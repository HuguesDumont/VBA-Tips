VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19110
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LabLeft As Single
Private LabTop  As Single
Private X0      As Single
Private Y0      As Single

Private Sub Label1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MsgBox "It works !"
End Sub

Private Sub Label1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   LabLeft = Label1.Left
   LabTop = Label1.Top
   X0 = X
   Y0 = Y
End Sub
   
Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 0 Then
        Exit Sub
    End If
    
    Label1.Left = Label1.Left + X - X0
    Label1.Top = Label1.Top + Y - Y0
End Sub
