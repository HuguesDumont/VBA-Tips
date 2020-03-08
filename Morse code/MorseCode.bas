Attribute VB_Name = "MorseCode"
Attribute VB_Description = "A simple module to convert text to morse code and morse code to text"
Option Explicit

'Matching table (2D Array)
Private trad(73, 1) As String

'Sub for testing
Sub test()
    Init
    Debug.Print MorseToText("... --- ...   ..   -. . . -..   .... . .-.. .--.")
    Debug.Print TextToMorse("SOS I NEED HELP")
    
    Debug.Print MorseToText("- . ... -   .-...   -.-. .- .-. .- -.-. - .-..- .-. . ...   ... .--. ..-.. -.-. .. .- ..- -..-   -.--. -.-.. .-   ...- .- ..- -   .-.. .   -.-. --- ..- .--.   ..--.. -.--.-")
    Debug.Print TextToMorse("Test & caractères spéciaux (ça vaut le coup ?)")
End Sub

'Function to convert morse code to text
Public Function MorseToText(codeMorse As String) As String
Attribute MorseToText.VB_Description = "Convert the morse code passed as a String parameter to text"
    Dim tmp() As String
    Dim tmpWord() As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    tmp = Split(codeMorse, "  ")
    
    For i = 0 To UBound(tmp)
        tmpWord = Split(tmp(i), " ")
        
        For j = 0 To UBound(tmpWord)
            For k = 0 To 72
                If (tmpWord(j) = trad(k, 1)) Then
                    MorseToText = MorseToText & UCase(trad(k, 0))
                    Exit For
                End If
            Next k
        Next j
        MorseToText = MorseToText & " "
    Next i
End Function

'Function to convert text to morse code
Public Function TextToMorse(text As String) As String
Attribute TextToMorse.VB_Description = "Convert the text passed as a String parameter to morse code"
    Dim i As Long, j As Long
    
    text = UCase(text)
    
    For i = 1 To Len(text)
        For j = 0 To 72
            If (Mid(text, i, 1) = trad(j, 0)) Then
                TextToMorse = TextToMorse & trad(j, 1) & " "
                Exit For
            End If
        Next j
    Next i
End Function

'Initialize the matching "table" (2D array)
Public Sub Init()
Attribute Init.VB_Description = "Initialize the matching ""table"" (2D array"
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
    
    trad(37, 0) = "."
    trad(37, 1) = ".-.-.-"
    
    trad(38, 0) = ","
    trad(38, 1) = "--..--"
    
    trad(39, 0) = "?"
    trad(39, 1) = "..--.."
    
    trad(40, 0) = "'"
    trad(40, 1) = ".----."
    
    trad(41, 0) = "!"
    trad(41, 1) = "-.-.-----."
    
    trad(42, 0) = "/"
    trad(42, 1) = "-..-."
    
    trad(43, 0) = "("
    trad(43, 1) = "-.--."
    
    trad(44, 0) = ")"
    trad(44, 1) = "-.--.-"
    
    trad(45, 0) = "&"
    trad(45, 1) = ".-..."
    
    trad(46, 0) = ":"
    trad(46, 1) = "---..."
    
    trad(47, 0) = ";"
    trad(47, 1) = "-.-.-."
    
    trad(48, 0) = "="
    trad(48, 1) = "-...-"
    
    trad(49, 0) = "+"
    trad(49, 1) = ".-.-."
    
    trad(50, 0) = "-"
    trad(50, 1) = "-....-"
    
    trad(51, 0) = "_"
    trad(51, 1) = "..--.-"
    
    trad(52, 0) = Chr(34)
    trad(52, 1) = ".-..-."
    
    trad(53, 0) = "$"
    trad(53, 1) = "...-..-"
    
    trad(54, 0) = "@"
    trad(54, 1) = ".--.-."
    
    trad(55, 0) = "Ä"
    trad(55, 1) = ".-.-"
    
    trad(56, 0) = "À"
    trad(56, 1) = ".--.-"
    
    trad(57, 0) = "Ç"
    trad(57, 1) = "-.-.."
    
    trad(58, 0) = "È"
    trad(58, 1) = ".-..-"
    
    trad(59, 0) = "É"
    trad(59, 1) = "..-.."
    
    trad(60, 0) = "Ñ"
    trad(60, 1) = "--.--"
    
    trad(61, 0) = "Ö"
    trad(61, 1) = "---."
    
    trad(62, 0) = "Ü"
    trad(62, 1) = "..--"
    
    trad(63, 0) = "Error"
    trad(63, 1) = "........"
    
    trad(64, 0) = "Start of transmission"
    trad(64, 1) = "-.-.-"
    
    trad(65, 0) = "End of transmission"
    trad(65, 1) = ".-.-."
    
    trad(66, 0) = "End of contact"
    trad(66, 1) = "...-.-"
    
    trad(67, 0) = "Understood"
    trad(67, 1) = "...-."
    
    trad(68, 0) = "Slower"
    trad(68, 1) = "....-.."
    
    trad(69, 0) = "Wrong signal"
    trad(69, 1) = ".-...-."
    
    trad(70, 0) = "Lighter"
    trad(70, 1) = ".-...-.."
    
    trad(71, 0) = "Darker"
    trad(71, 1) = ".--..--."
    
    trad(72, 0) = "SOS"
    trad(72, 1) = "...---..."
End Sub


































































