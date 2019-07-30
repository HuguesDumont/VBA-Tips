Attribute VB_Name = "StringInt"
Attribute VB_Description = "Functions for String with Integer type"
Option Explicit

'Function to check if string is an integer
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function IsInteger(ByVal value As String) As Boolean
Attribute IsInteger.VB_Description = "Function to check if string is an integer\r\nNeed to activate the reference ""Microsoft VBScrpt Regular Expressions 5.5"""
    Dim reg As New VBScript_RegExp_55.RegExp
    Const INT_MIN As Integer = -32768
    Const INT_MAX As Integer = 32767
    
    reg.Pattern = "^(-)?(\d)+$"
    IsInteger = False
    If reg.test(value) Then
        On Error GoTo capacityOverflow
        If CInt(value) >= INT_MIN And CInt(value) <= INT_MAX Then
            IsInteger = True
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
Public Function IsIntPos(ByVal value As String) As Boolean
Attribute IsIntPos.VB_Description = "Function to check if string is an integer over zero\r\nIt uses the ""isInteger"" function implemented before"
    IsIntPos = False
    If IsInteger(value) Then
        If CInt(value) > 0 Then
            IsIntPos = True
        End If
    End If
End Function

'Function to check if string is integer below zero
'It uses the "isInteger" function implemented before
Public Function IsIntNeg(ByVal value As String) As Boolean
Attribute IsIntNeg.VB_Description = "Function to check if string is integer below zero\r\nIt uses the ""isInteger"" function implemented before"
    IsIntNeg = False
    If IsInteger(value) Then
        If CInt(value) < 0 Then
            IsIntNeg = True
        End If
    End If
End Function

'Function to check if string is zero (integer 0)
'It uses the "isInteger" function implemented before
Public Function IsZero(ByVal value As String) As Boolean
Attribute IsZero.VB_Description = "Function to check if string is zero (integer 0)\r\nIt uses the ""isInteger"" function implemented before"
    IsZero = False
    If IsInteger(value) Then
        If CInt(value) = 0 Then
            IsZero = True
        End If
    End If
End Function

'Function to check if string is integer above or equal zero
'It uses the "isInteger" function implemented before
Public Function IsIntPosOrZero(ByVal value As String) As Boolean
Attribute IsIntPosOrZero.VB_Description = "Function to check if string is integer above or equal zero\r\nIt uses the ""isInteger"" function implement before"
    IsIntPosOrZero = False
    If IsInteger(value) Then
        If CInt(value) >= 0 Then
            IsIntPosOrZero = True
        End If
    End If
End Function

'Function to check if string is integer above or equal zero
'It uses the "isInteger" function implemented before
Public Function IsIntNegOrZero(ByVal value As String) As Boolean
Attribute IsIntNegOrZero.VB_Description = "Function to check if string is integer above or equal zero\r\nIt uses the ""isInteger"" function implement before"
    IsIntNegOrZero = False
    If IsInteger(value) Then
        If CInt(value) <= 0 Then
            IsIntNegOrZero = True
        End If
    End If
End Function
