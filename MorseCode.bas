Attribute VB_Name = "Module1"
Option Explicit

Private trad(37, 1) As String

Sub test()
    init
    Debug.Print MorseToText("... --- ...   ..   -. . . -..   .... . .-.. .--.")
    Debug.Print TextToMorse("SOS I NEED HELP")
End Sub

'Function to convert morse code to text
Public Function MorseToText(codeMorse As String) As String
    Dim tmp() As String, tmpWord() As String
    Dim i As Long, j As Long, k As Long
    
    tmp = Split(codeMorse, "  ")
    For i = 0 To UBound(tmp)
        tmpWord = Split(tmp(i), " ")
        
        For j = 0 To UBound(tmpWord)
            For k = 0 To 36
                If tmpWord(j) = trad(k, 1) Then
                    MorseToText = MorseToText & trad(k, 0)
                    Exit For
                End If
            Next k
        Next j
        MorseToText = MorseToText & " "
    Next i
End Function

'Function to convert text to morse code
Public Function TextToMorse(text As String) As String
    Dim i As Long, j As Long
    
    text = UCase(text)
    For i = 1 To Len(text)
        For j = 0 To 36
            If Mid(text, i, 1) = trad(j, 0) Then
                TextToMorse = TextToMorse & trad(j, 1) & " "
                Exit For
            End If
        Next j
    Next i
End Function

Public Sub init()
    trad(0, 0) = "A"
    trad(0, 1) = ".-"
    
    trad(1, 0) = "B"
    trad(1, 1) = "-..."
    
    trad(2, 0) = "C"
    trad(2, 1) = "-.-."
    
    trad(3, 0) = "D"
    trad(3, 1) = "-.."
    
    trad(4, 0) = "E"
    trad(4, 1) = "."
    
    trad(5, 0) = "F"
    trad(5, 1) = "..-."
    
    trad(6, 0) = "G"
    trad(6, 1) = "--."
    
    trad(7, 0) = "H"
    trad(7, 1) = "...."
    
    trad(8, 0) = "I"
    trad(8, 1) = ".."
    
    trad(9, 0) = "J"
    trad(9, 1) = ".---"
    
    trad(10, 0) = "K"
    trad(10, 1) = "-.-"
    
    trad(11, 0) = "L"
    trad(11, 1) = ".-.."
    
    trad(12, 0) = "M"
    trad(12, 1) = "--"
    
    trad(13, 0) = "N"
    trad(13, 1) = "-."
    
    trad(14, 0) = "O"
    trad(14, 1) = "---"
    
    trad(15, 0) = "P"
    trad(15, 1) = ".--."
    
    trad(16, 0) = "Q"
    trad(16, 1) = "--.-"
    
    trad(17, 0) = "R"
    trad(17, 1) = ".-."
    
    trad(18, 0) = "S"
    trad(18, 1) = "..."
    
    trad(19, 0) = "T"
    trad(19, 1) = "-"
    
    trad(20, 0) = "U"
    trad(20, 1) = "..-"
    
    trad(21, 0) = "V"
    trad(21, 1) = "...-"
    
    trad(22, 0) = "W"
    trad(22, 1) = ".--"
    
    trad(23, 0) = "X"
    trad(23, 1) = "-..-"
    
    trad(24, 0) = "Y"
    trad(24, 1) = "-.--"
    
    trad(25, 0) = "Z"
    trad(25, 1) = "--.."
    
    trad(26, 0) = "1"
    trad(26, 1) = ".----"
    
    trad(27, 0) = "2"
    trad(27, 1) = "..---"
    
    trad(28, 0) = "3"
    trad(28, 1) = "...--"
    
    trad(29, 0) = "4"
    trad(29, 1) = "....-"
    
    trad(30, 0) = "5"
    trad(30, 1) = "....."
    
    trad(31, 0) = "6"
    trad(31, 1) = "-...."
    
    trad(32, 0) = "7"
    trad(32, 1) = "--..."
    
    trad(33, 0) = "8"
    trad(33, 1) = "---.."
    
    trad(34, 0) = "9"
    trad(34, 1) = "----."
    
    trad(35, 0) = "0"
    trad(35, 1) = "-----"

    trad(36, 0) = " "
    trad(36, 1) = " "
End Sub
