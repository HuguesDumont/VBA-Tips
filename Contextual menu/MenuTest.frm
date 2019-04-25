VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MenuTest 
   Caption         =   "UserForm1"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "MenuTest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MenuTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub TextBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Select Case Button
        Case XlMouseButton.xlSecondaryButton
            CommandBars("Contextual").ShowPopup
    End Select
End Sub
 
Private Sub TextBox2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Select Case Button
        Case XlMouseButton.xlSecondaryButton
            CommandBars("Contextual").ShowPopup
    End Select
End Sub

Private Sub UserForm_Initialize()
    Call ContextualMenu.CreateContextual
    Set ContextualMenu.param = Me
End Sub
