Attribute VB_Name = "Others"
Attribute VB_Description = "Divers functions like conversion between roman and arabic numerals or temperature conversion"
Option Explicit

'Function to check string for balanced brackets
Public Function checkBrackets(str As String) As Boolean
Attribute checkBrackets.VB_Description = "Function to check string for balanced brackets"
    Dim i As Long
    Dim Depth As Integer
    Dim ch As String * 1
     
    Depth = 0
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        If ch = "[" Then
            Depth = Depth + 1
        End If
        
        If ch = "]" Then
            If Depth = 0 Then
                checkBrackets = False
                Exit Function
            Else
                Depth = Depth - 1
            End If
        End If
    Next i
    checkBrackets = (Depth = 0)
End Function

'Function to convert arabic numerals to romans (FR : nombres romains)
Public Function arabicToRomans(arabic As Long) As String
Attribute arabicToRomans.VB_Description = "Function to convert arabic numerals to romans (FR : nombres romains)"
    Const hundreds = ",C,CC,CCC,CD,D,DC,DCC,DCCC,CM"
    Const tenths = ",X,XX,XXX,XL,L,LX,LXX,LXXX,XC"
    Const units = ",I,II,III,IV,V,VI,VII,VIII,IX"
    
    If arabic < 0 Then
        arabicToRomans = "-"
        arabic = -arabic
    End If
    
    arabicToRomans = arabicToRomans & String(arabic \ 1000, "M")
    
    arabic = arabic Mod 1000
    arabicToRomans = arabicToRomans & Split(hundreds, ",")(arabic \ 100)
    
    arabic = arabic Mod 100
    arabicToRomans = arabicToRomans & Split(tenths, ",")(arabic \ 10)
    
    arabic = arabic Mod 10
    arabicToRomans = arabicToRomans & Split(units, ",")(arabic)
End Function

'Function to convert romans (FR : nombres romains) to arabic numerals
'If string isn't a valid roman numeral then return -1
Public Function romanToArabic(roman As String) As Long
Attribute romanToArabic.VB_Description = "Function to convert romans (FR : nombres romains) to arabic numerals\r\nIf string isn't a valid roman numeral then return -1"
    Dim i As Long, unit As Long, oldUnit As Long
    
    oldUnit = 1000
    
    roman = UCase(roman)
    For i = 1 To Len(roman)
        Select Case Mid(roman, i, 1)
            Case "I":  unit = 1
            Case "V":  unit = 5
            Case "X":  unit = 10
            Case "L":  unit = 50
            Case "C":  unit = 100
            Case "D":  unit = 500
            Case "M":  unit = 1000
            Case Else: 'invalid roman string because invalid character is detected
                romanToArabic = -1
                Exit Function
        End Select
        If unit > oldUnit Then
            romanToArabic = romanToArabic - 2 * oldUnit
        End If
        romanToArabic = romanToArabic + unit
        oldUnit = unit
    Next i
End Function

'Function to convert temperature between scales (Kelvin, Celsius, Fahreneit
'To chose source and destination scales :
'   0 : Kelvin
'   1 : Celsius
'   2 : Fahrenheit
Public Function temperatureConversion(Temp As Double, source As Integer, dest As Integer) As Single
Attribute temperatureConversion.VB_Description = "'Function to convert temperature between scales (Kelvin, Celsius, Fahreneit\r\nTo chose source and destination scales :\r\n   0 : Kelvin\r\n   1 : Celsius\r\n   2 : Fahrenheit"
    If source = dest Then
        temperatureConversion = Temp
    ElseIf dest = 0 Then
        If source = 1 Then
            temperatureConversion = Temp + 273.15
        Else
            temperatureConversion = (Temp - 32) * (5 / 9) + 273.15
        End If
    ElseIf dest = 1 Then
        If source = 0 Then
            temperatureConversion = Temp - 273.15
        Else
            temperatureConversion = (Temp - 32) * (5 / 9)
        End If
    Else
        If source = 0 Then
            temperatureConversion = (Temp - 273.15) * (9 / 5) + 32
        Else
            temperatureConversion = Temp * (9 / 5) + 32
        End If
    End If
End Function
