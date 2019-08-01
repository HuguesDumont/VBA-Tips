Attribute VB_Name = "BaseConversion"
Attribute VB_Description = "Functions for conversion between bases"
Option Explicit

'Function to convert from any base to decimal
'Base limit is 35 (9 digits (0 is excluded) + 26 letters)
Public Function BaseToDecimal(ByVal srcValue As String, ByVal srcBase As Integer) As Long
    Dim i As Long, digitValue As Long
    Dim strDigits As String

    strDigits = Left("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", srcBase)

    If srcBase < 2 Or srcBase > 36 Then
        MsgBox "Cannot convert from base over 36 or below 2.", vbOKOnly + vbExclamation, "Base error"
    End If

    For i = 1 To Len(srcValue)
        digitValue = InStr(1, strDigits, Mid(srcValue, i, 1), vbTextCompare) - 1
        If digitValue < 0 Then
            MsgBox "Unvalid charachter found in source value", vbOKOnly + vbExclamation, "Unvalid value"
        End If
        BaseToDecimal = BaseToDecimal * srcBase + digitValue
    Next i
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

'Function to convert binary string to decimal number
'Need the Microsoft VBScript Regular Expressions 5.5 reference to test if string is a binary string
'Return -1 if error occurred
Public Function BinaryToDecimal(ByVal strBin As String) As Long
    Dim x As Integer
    Dim reg As New VBScript_RegExp_55.RegExp
    Dim strPattern As String

    reg.Pattern = "^((0|1)*)$"
    If Not reg.test(strBin) Then
        MsgBox "Cannot convert to decimal value because string isn't a binary number.", vbOKOnly + vbCritical, "Invalid string"
        BinaryToDecimal = -1
        Exit Function
    End If

    On Error GoTo cannotConvert

    For x = 0 To Len(strBin) - 1
        BinaryToDecimal = CDec(BinaryToDecimal) + val(Mid(strBin, Len(strBin) - x, 1)) * 2 ^ x
    Next x
    Exit Function

cannotConvert:
    If Err.number = 13 Then
        MsgBox "Cannot convert to decimal. Please check that your value is a valid binary string.", vbCritical + vbOKOnly, "Invalid string"
    ElseIf Err.number = 6 Then
        MsgBox "Cannot convert to decimal. Value out of bound.", vbCritical + vbOKOnly, "Value out of bound"
    Else
        MsgBox "Cannot convert to decimal due to an unknown error.", vbCritical + vbOKOnly, "Unknown error"
    End If
    BinaryToDecimal = -1
End Function

'Function to convert from any base to any other base
'Using decimal as temporary conversion
'Using baseToDecimal and decimalToBase functions
Public Function ConvertBase(ByVal value As String, ByVal srcBase As Integer, ByVal destBase As Integer) As String
    ConvertBase = DecimalToBase(BaseToDecimal(value, srcBase), destBase)
End Function

'Function to convert a decimal value to any other base
'Base limit is 35 (9 digits (0 is excluded) + 26 letters)
Public Function DecimalToBase(ByVal srcValue As String, ByVal destBase As Integer) As String
    Dim valueRest As Long, toDivide As Long
    Dim charRest As String

    On Error GoTo cannotConvert

    toDivide = val(srcValue)
    DecimalToBase = ""

    If destBase < 2 Or destBase > 36 Then
        MsgBox "Cannot convert to base over 36 or below 2.", vbOKOnly + vbExclamation, "Base error"
    End If

    While toDivide > 0
        valueRest = toDivide - Int(toDivide / destBase) * destBase
        charRest = Mid("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", valueRest + 1, 1)
        DecimalToBase = Trim(charRest) & DecimalToBase
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

'Function to convert positive long number to binary string
'Return empty string if error occurred
Public Function DecimalToBinary(ByVal strDec As Long) As String
    Dim isNeg As Boolean

    On Error GoTo cannotConvert

    isNeg = (strDec < 0)
    If isNeg Then
        strDec = strDec * -1
    End If

    While strDec <> 0
        DecimalToBinary = Format(strDec - 2 * Int(strDec / 2)) & DecimalToBinary
        strDec = Int(strDec / 2)
    Wend

    If isNeg Then
        DecimalToBinary = "-" & DecimalToBinary
    End If
    Exit Function

cannotConvert:
    If Err.number = 6 Then
        MsgBox "Cannot convert to binary. Integer out of bound.", vbCritical + vbOKOnly, "Integer out of bound"
    Else
        MsgBox "Cannot convert to binary due to an unknown error.", vbCritical + vbOKOnly, "Unknown error"
    End If
    DecimalToBinary = ""
End Function
