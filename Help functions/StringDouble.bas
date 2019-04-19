Attribute VB_Name = "StringDouble"
Option Explicit

'Function to check if string is of Double type
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function isDouble(value As String) As Boolean
Attribute isDouble.VB_Description = "Function to check if string is of Double type\r\nNeed to activate the reference ""Microsoft VBScrpt Regular Expressions 5.5"""
    Dim reg As New VBScript_RegExp_55.RegExp
    Const Double_MIN As Double = -4.94065645841247E-324
    Const Double_MAX As Double = 1.79769313486231E+308
    
    reg.Pattern = "^(-)?(\d)+(((\.)|(,))(\d)+)?$"
    isDouble = False
    If reg.test(value) Then
        On Error GoTo capacityOverflow
        If ((CDbl(value) >= Double_MIN) And (CDbl(value) <= Double_MAX)) Then
            isDouble = True
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
Public Function isDoublePos(value As String) As Boolean
Attribute isDoublePos.VB_Description = "Function to check if string is a Double over zero\r\nUsing ""isDouble"" function"
    isDoublePos = False
    If isDouble(value) Then
        If CLng(value) > 0 Then
            isDoublePos = True
        End If
    End If
End Function

'Function to check if string is Double below zero
'Using "isDouble" function
Public Function isDoubleNeg(value As String) As Boolean
Attribute isDoubleNeg.VB_Description = "Function to check if string is Double below zero\r\nUsing ""isDouble"" function"
    isDoubleNeg = False
    If isDouble(value) Then
        If CLng(value) < 0 Then
            isDoubleNeg = True
        End If
    End If
End Function

'Function to check if string is zero (integer 0)
'Using "isDouble" function
Public Function isZero(value As String) As Boolean
Attribute isZero.VB_Description = "Function to check if string is zero (integer 0)\r\nUsing ""isDouble"" function"
    isZero = False
    If isDouble(value) Then
        If CLng(value) = 0 Then
            isZero = True
        End If
    End If
End Function

'Function to check if string is Double above or equal zero
'Using "isDouble" function
Public Function isDoublePosOrZero(value As String) As Boolean
Attribute isDoublePosOrZero.VB_Description = "Function to check if string is Double above or equal zero\r\nUsing ""isDouble"" function"
    isDoublePosOrZero = False
    If isDouble(value) Then
        If CLng(value) >= 0 Then
            isDoublePosOrZero = True
        End If
    End If
End Function

'Function to check if string is Double above or equal zero
'Using "isDouble" function
Public Function isDoubleNegOrZero(value As String) As Boolean
Attribute isDoubleNegOrZero.VB_Description = "Function to check if string is Double above or equal zero\r\nUsing ""isDouble"" function"
    isDoubleNegOrZero = False
    If isDouble(value) Then
        If CLng(value) <= 0 Then
            isDoubleNegOrZero = True
        End If
    End If
End Function
