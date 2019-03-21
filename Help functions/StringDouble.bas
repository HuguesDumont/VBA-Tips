Attribute VB_Name = "StringDouble"
Option Explicit

'Function to check if string is of Double type
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function isDouble(value As String) As Boolean
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
'It uses the "isDouble" function implemented before
Public Function isDoublePos(value As String) As Boolean
    isDoublePos = False
    If isDouble(value) Then
        If CLng(value) > 0 Then
            isDoublePos = True
        End If
    End If
End Function

'Function to check if string is Double below zero
'It uses the "isDouble" function implemented before
Public Function isDoubleNeg(value As String) As Boolean
    isDoubleNeg = False
    If isDouble(value) Then
        If CLng(value) < 0 Then
            isDoubleNeg = True
        End If
    End If
End Function

'Function to check if string is zero (integer 0)
'It uses the "isDouble" function implemented before
Public Function isZero(value As String) As Boolean
    isZero = False
    If isDouble(value) Then
        If CLng(value) = 0 Then
            isZero = True
        End If
    End If
End Function

'Function to check if string is Double above or equal zero
'It uses the "isDouble" function implement before
Public Function isDoublePosOrZero(value As String) As Boolean
    isDoublePosOrZero = False
    If isDouble(value) Then
        If CLng(value) >= 0 Then
            isDoublePosOrZero = True
        End If
    End If
End Function

'Function to check if string is Double above or equal zero
'It uses the "isDouble" function implement before
Public Function isDoubleNegOrZero(value As String) As Boolean
    isDoubleNegOrZero = False
    If isDouble(value) Then
        If CLng(value) <= 0 Then
            isDoubleNegOrZero = True
        End If
    End If
End Function


