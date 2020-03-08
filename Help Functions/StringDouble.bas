Attribute VB_Name = "StringDouble"
Attribute VB_Description = "Functions for String with Double type\n"
Option Explicit

'Function to check if string is of Double type
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function IsDouble(ByVal value As String) As Boolean
    Const DOUBLE_MIN                    As Double = -4.94065645841247E-324
    Const DOUBLE_MAX                    As Double = 1.79769313486231E+308
    
    Dim reg                             As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^(-)?(\d)+(((\.)|(,))(\d)+)?$"
    IsDouble = False
    
    If reg.test(value) Then
        On Error GoTo capacityOverflow
        IsDouble = ((CDbl(value) >= DOUBLE_MIN) And (CDbl(value) <= DOUBLE_MAX))
    End If
    
    Set reg = Nothing
    Exit Function
    
capacityOverflow:
    MsgBox "Value is single but over 1.79769313486231E+308 or lower than -4.94065645841247E-324" & Chr(13) & "Can't be converted to the Double type in vba!", _
            vbOKOnly + vbCritical, "Capacity overflow !"
    Set reg = Nothing
End Function

'Function to check if string is Double below zero
'Using "isDouble" function
Public Function IsDoubleNeg(ByVal value As String) As Boolean
    IsDoubleNeg = False
    
    If IsDouble(value) Then
        IsDoubleNeg = (CLng(value) < 0)
    End If
End Function

'Function to check if string is Double above or equal zero
'Using "isDouble" function
Public Function IsDoubleNegOrZero(ByVal value As String) As Boolean
    IsDoubleNegOrZero = False
    
    If IsDouble(value) Then
        IsDoubleNegOrZero = (CLng(value) <= 0)
    End If
End Function

'Function to check if string is a Double over zero
'Using "isDouble" function
Public Function IsDoublePos(ByVal value As String) As Boolean
    IsDoublePos = False
    
    If IsDouble(value) Then
        IsDoublePos = (CLng(value) > 0)
    End If
End Function

'Function to check if string is Double above or equal zero
'Using "isDouble" function
Public Function IsDoublePosOrZero(ByVal value As String) As Boolean
    IsDoublePosOrZero = False
    
    If IsDouble(value) Then
        IsDoublePosOrZero = (CLng(value) >= 0)
    End If
End Function

'Function to check if string is zero (integer 0)
'Using "isDouble" function
Public Function IsZero(ByVal value As String) As Boolean
    IsZero = False
    
    If IsDouble(value) Then
        IsZero = (CLng(value) = 0)
    End If
End Function
