Attribute VB_Name = "StringInt"
Attribute VB_Description = "Functions for String with Integer type"
Option Explicit

'Function to check if string is an integer
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function isInteger(value As String) As Boolean
Attribute isInteger.VB_Description = "Function to check if string is an integer\r\nNeed to activate the reference ""Microsoft VBScrpt Regular Expressions 5.5"""
    Dim reg As New VBScript_RegExp_55.RegExp
    Const INT_MIN As Integer = -32768
    Const INT_MAX As Integer = 32767
    
    reg.Pattern = "^(-)?(\d)+$"
    isInteger = False
    If reg.test(value) Then
        On Error GoTo capacityOverflow
        If CInt(value) >= INT_MIN And CInt(value) <= INT_MAX Then
            isInteger = True
        End If
    End If
    Set reg = Nothing
    Exit Function
capacityOverflow:
    MsgBox "Value is integer but over 32 767 or lower than -32 768" & Chr(13) & _
        "Can't be converted to the integer type in vba (might be able to convert to long type)!", _
            vbOKOnly + vbCritical, "Capacity overflow !"
    Set reg = Nothing
End Function

'Function to check if string is an integer over zero
'It uses the "isInteger" function implemented before
Public Function isIntPos(value As String) As Boolean
Attribute isIntPos.VB_Description = "Function to check if string is an integer over zero\r\nIt uses the ""isInteger"" function implemented before"
    isIntPos = False
    If isInteger(value) Then
        If CInt(value) > 0 Then
            isIntPos = True
        End If
    End If
End Function

'Function to check if string is integer below zero
'It uses the "isInteger" function implemented before
Public Function isIntNeg(value As String) As Boolean
Attribute isIntNeg.VB_Description = "Function to check if string is integer below zero\r\nIt uses the ""isInteger"" function implemented before"
    isIntNeg = False
    If isInteger(value) Then
        If CInt(value) < 0 Then
            isIntNeg = True
        End If
    End If
End Function

'Function to check if string is zero (integer 0)
'It uses the "isInteger" function implemented before
Public Function isZero(value As String) As Boolean
Attribute isZero.VB_Description = "Function to check if string is zero (integer 0)\r\nIt uses the ""isInteger"" function implemented before"
    isZero = False
    If isInteger(value) Then
        If CInt(value) = 0 Then
            isZero = True
        End If
    End If
End Function

'Function to check if string is integer above or equal zero
'It uses the "isInteger" function implement before
Public Function isIntPosOrZero(value As String) As Boolean
Attribute isIntPosOrZero.VB_Description = "Function to check if string is integer above or equal zero\r\nIt uses the ""isInteger"" function implement before"
    isIntPosOrZero = False
    If isInteger(value) Then
        If CInt(value) >= 0 Then
            isIntPosOrZero = True
        End If
    End If
End Function

'Function to check if string is integer above or equal zero
'It uses the "isInteger" function implement before
Public Function isIntNegOrZero(value As String) As Boolean
Attribute isIntNegOrZero.VB_Description = "Function to check if string is integer above or equal zero\r\nIt uses the ""isInteger"" function implement before"
    isIntNegOrZero = False
    If isInteger(value) Then
        If CInt(value) <= 0 Then
            isIntNegOrZero = True
        End If
    End If
End Function
