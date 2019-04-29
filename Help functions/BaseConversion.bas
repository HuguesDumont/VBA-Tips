Attribute VB_Name = "BaseConversion"
Attribute VB_Description = "Functions for conversion between bases"
Option Explicit

'Function to convert positive long number to binary string
'Return empty string if error occurred
Public Function decimalToBinary(strDec As Long) As String
    Dim isNeg As Boolean
    On Error GoTo cannotConvert
    
    isNeg = (strDec < 0)
    If isNeg Then
        strDec = strDec * -1
    End If
    
    While strDec <> 0
        decimalToBinary = Format(strDec - 2 * Int(strDec / 2)) & decimalToBinary
        strDec = Int(strDec / 2)
    Wend
    
    If isNeg Then
        decimalToBinary = "-" & decimalToBinary
    End If
    
    Exit Function
    
cannotConvert:
    If Err.number = 6 Then
        MsgBox "Cannot convert to binary. Integer out of bound.", vbCritical + vbOKOnly, "Integer out of bound"
    Else
        MsgBox "Cannot convert to binary due to an unknown error.", vbCritical + vbOKOnly, "Unknown error"
    End If
    decimalToBinary = ""
End Function

'Function to convert binary string to decimal number
'Need the Microsoft VBScript Regular Expressions 5.5 reference to test if string is a binary string
'Return -1 if error occurred
Public Function binaryToDecimal(strBin As String) As Long
Attribute binaryToDecimal.VB_Description = "Function to convert binary string to decimal number\r\nNeed the Microsoft VBScript Regular Expressions 5.5 reference to test if string is a binary string\r\nReturn -1 if error occurred"
    Dim x As Integer
    Dim reg As New VBScript_RegExp_55.RegExp
    Dim strPattern As String

    reg.Pattern = "^((0|1)*)$"
    If Not reg.test(strBin) Then
        MsgBox "Cannot convert to decimal value because string isn't a binary number.", vbOKOnly + vbCritical, "Invalid string"
        binaryToDecimal = -1
        Exit Function
    End If
    
    On Error GoTo cannotConvert
    
    For x = 0 To Len(strBin) - 1
        binaryToDecimal = CDec(binaryToDecimal) + val(Mid(strBin, Len(strBin) - x, 1)) * 2 ^ x
    Next
    
    Exit Function
    
cannotConvert:
    If Err.number = 13 Then
        MsgBox "Cannot convert to decimal. Please check that your value is a valid binary string.", vbCritical + vbOKOnly, "Invalid string"
    ElseIf Err.number = 6 Then
        MsgBox "Cannot convert to decimal. Value out of bound.", vbCritical + vbOKOnly, "Value out of bound"
    Else
        MsgBox "Cannot convert to decimal due to an unknown error.", vbCritical + vbOKOnly, "Unknown error"
    End If
    binaryToDecimal = -1
End Function

'Function to convert a decimal value to any other base
'Base limit is 35 (9 digits (0 is excluded) + 26 letters)
Public Function decimalToBase(srcValue As String, destBase As Integer) As String
Attribute decimalToBase.VB_Description = "Function to convert a decimal value to any other base\r\nBase limit is 35 (9 digits (0 is excluded) + 26 letters)"
    Dim valueRest As Long
    Dim charRest As String
    Dim toDivide As Long
    
    On Error GoTo cannotConvert
    
    toDivide = val(srcValue)
    decimalToBase = ""
    
    If destBase < 2 Or destBase > 36 Then
        MsgBox "Cannot convert to base over 36 or below 2.", vbOKOnly + vbExclamation, "Base error"
    End If
    
    While toDivide > 0
        valueRest = toDivide - Int(toDivide / destBase) * destBase
        charRest = Mid("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", valueRest + 1, 1)
        decimalToBase = Trim(charRest) & decimalToBase
        toDivide = Int(toDivide / destBase)
    Wend
    
    Exit Function

cannotConvert:
    If Err.number = 13 Then
        MsgBox "Cannot convert to decimal. Please check that your value is a valid binary string.", vbCritical + vbOKOnly, "Invalid string"
    ElseIf Err.number = 6 Then
        MsgBox "Cannot convert to decimal. Value out of bound.", vbCritical + vbOKOnly, "Value out of bound"
    Else
        MsgBox "Cannot convert to decimal due to an unknown error.", vbCritical + vbOKOnly, "Unknown error"
    End If
End Function

'Function to convert from any base to decimal
'Base limit is 35 (9 digits (0 is excluded) + 26 letters)
Function baseToDecimal(srcValue As String, srcBase As Integer) As Long
Attribute baseToDecimal.VB_Description = "Function to convert from any base to decimal\r\nBase limit is 35 (9 digits (0 is excluded) + 26 letters)"
    Dim i As Long
    Dim strDigits As String
    Dim digitValue As Long
    
    strDigits = Left("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", srcBase)
    
    If srcBase < 2 Or srcBase > 36 Then
        MsgBox "Cannot convert from base over 36 or below 2.", vbOKOnly + vbExclamation, "Base error"
    End If
        
    For i = 1 To Len(srcValue)
        digitValue = InStr(1, strDigits, Mid(srcValue, i, 1), vbTextCompare) - 1
        If digitValue < 0 Then
            MsgBox "Unvalid charachter found in source value", vbOKOnly + vbExclamation, "Unvalid value"
        End If
        baseToDecimal = baseToDecimal * srcBase + digitValue
    Next
    
    Exit Function

cannotConvert:
    If Err.number = 13 Then
        MsgBox "Cannot convert to decimal. Please check that your value is a valid binary string.", vbCritical + vbOKOnly, "Invalid string"
    ElseIf Err.number = 6 Then
        MsgBox "Cannot convert to decimal. Value out of bound.", vbCritical + vbOKOnly, "Value out of bound"
    Else
        MsgBox "Cannot convert to decimal due to an unknown error.", vbCritical + vbOKOnly, "Unknown error"
    End If
End Function

'Function to convert from any base to any other base
'Using decimal as temporary conversion
'Using baseToDecimal and decimalToBase functions
Public Function convertBase(value As String, srcBase As Integer, destBase As Integer) As String
Attribute convertBase.VB_Description = "Function to convert from any base to any other base\r\nUsing decimal as temporary conversion\r\nUsing baseToDecimal and decimalToBase functions"
    convertBase = decimalToBase(baseToDecimal(value, srcBase), destBase)
End Function
