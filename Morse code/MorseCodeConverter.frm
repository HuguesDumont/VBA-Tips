VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MorseCodeConverter 
   Caption         =   "Morse code converter"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20625
   OleObjectBlob   =   "MorseCodeConverter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MorseCodeConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A simple screen to convert your text to morse code and inversely"
Option Explicit

Private Sub UserForm_Initialize()
    Call MorseCode.init
End Sub

Private Sub textToMorseButton_Click()
    morseValue.text = MorseCode.TextToMorse(textValue.text)
End Sub

Private Sub morseToTextButton_Click()
    textValue.text = MorseCode.MorseToText(morseValue.text)
End Sub
