Attribute VB_Name = "BaseConversion"
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
        binaryToDecimal = CDec(binaryToDecimal) + Val(Mid(strBin, Len(strBin) - x, 1)) * 2 ^ x
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
