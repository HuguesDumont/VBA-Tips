Attribute VB_Name = "StringDouble"
Attribute VB_Description = "Functions for String with Double type\n"
Option Explicit

'Function to check if string is of Double type
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function IsDouble(ByVal value As String) As Boolean
Attribute IsDouble.VB_Description = "Function to check if string is of Double type\r\nNeed to activate the reference ""Microsoft VBScrpt Regular Expressions 5.5"""
    Dim reg As New VBScript_RegExp_55.RegExp
    Const Double_MIN As Double = -4.94065645841247E-324
    Const Double_MAX As Double = 1.79769313486231E+308
    
    reg.Pattern = "^(-)?(\d)+(((\.)|(,))(\d)+)?$"
    IsDouble = False
    If reg.test(value) Then
        On Error GoTo capacityOverflow
        If ((CDbl(value) >= Double_MIN) And (CDbl(value) <= Double_MAX)) Then
            IsDouble = True
        End If
    End If
    Set reg = Nothing
    Exit Function
capacityOverflow:
    MsgBox "Value is single but over 1.79769313486231E+308 or lower than -4.94065645841247E-324" & Chr(13) & _
        "Can't be converted to the Double type in vba!", _
        vbOKOnly + vbCritical, "Capacity overflow !"
    Set reg = Nothing
End Function

'Function to check if string is a Double over zero
'Using "isDouble" function
Public Function IsDoublePos(ByVal value As String) As Boolean
Attribute IsDoublePos.VB_Description = "Function to check if string is a Double over zero\r\nUsing ""isDouble"" function"
    IsDoublePos = False
    If IsDouble(value) Then
        If CLng(value) > 0 Then
            IsDoublePos = True
        End If
    End If
End Function

'Function to check if string is Double below zero
'Using "isDouble" function
Public Function IsDoubleNeg(ByVal value As String) As Boolean
Attribute IsDoubleNeg.VB_Description = "Function to check if string is Double below zero\r\nUsing ""isDouble"" function"
    IsDoubleNeg = False
    If IsDouble(value) Then
        If CLng(value) < 0 Then
            IsDoubleNeg = True
        End If
    End If
End Function

'Function to check if string is zero (integer 0)
'Using "isDouble" function
Public Function IsZero(ByVal value As String) As Boolean
Attribute IsZero.VB_Description = "Function to check if string is zero (integer 0)\r\nUsing ""isDouble"" function"
    IsZero = False
    If IsDouble(value) Then
        If CLng(value) = 0 Then
            IsZero = True
        End If
    End If
End Function

'Function to check if string is Double above or equal zero
'Using "isDouble" function
Public Function IsDoublePosOrZero(ByVal value As String) As Boolean
Attribute IsDoublePosOrZero.VB_Description = "Function to check if string is Double above or equal zero\r\nUsing ""isDouble"" function"
    IsDoublePosOrZero = False
    If IsDouble(value) Then
        If CLng(value) >= 0 Then
            IsDoublePosOrZero = True
        End If
    End If
End Function

'Function to check if string is Double above or equal zero
'Using "isDouble" function
Public Function IsDoubleNegOrZero(ByVal value As String) As Boolean
Attribute IsDoubleNegOrZero.VB_Description = "Function to check if string is Double above or equal zero\r\nUsing ""isDouble"" function"
    IsDoubleNegOrZero = False
    If IsDouble(value) Then
        If CLng(value) <= 0 Then
            IsDoubleNegOrZero = True
        End If
    End If
End Function
