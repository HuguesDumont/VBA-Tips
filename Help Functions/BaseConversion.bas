Attribute VB_Name = "BaseConversion"
Attribute VB_Description = "Functions for conversion between bases"
Option Explicit

'Convert from any base to decimal
'Base limit is 35 (9 digits (0 is excluded) + 26 letters)
'Parameters :
'- srcValue : source value in the other base
'- srvBase  : source base (number of units in source base)
'Returns : the value (Long) in decimal base (base 10)
Public Function BaseToDecimal(ByVal srcValue As String, ByVal srcBase As Integer) As Long
    Dim i                               As Long
    Dim digitValue                      As Long
    Dim strDigits                       As String
    
    On Error GoTo cannotConvert
    
    strDigits = Left("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", srcBase)
    
    If ((srcBase < 2) Or (srcBase > 36)) Then
        MsgBox "Cannot convert from base over 36 or below 2.", vbOKOnly + vbExclamation, "Base error"
    End If
    
    For i = 1 To Len(srcValue)
        digitValue = InStr(1, strDigits, Mid(srcValue, i, 1), vbTextCompare) - 1
        
        If (digitValue < 0) Then
            MsgBox "Unvalid charachter found in source value", vbOKOnly + vbExclamation, "Unvalid value"
        End If
        
        BaseToDecimal = BaseToDecimal * srcBase + digitValue
    Next i
    
    Exit Function
    
cannotConvert:
    If (Err.number = 13) Then
        Err.Raise 1013, "BaseToDecimal", "Cannot convert to decimal. Please check that your value is a valid number based string."
    ElseIf (Err.number = 6) Then
        Err.Raise 1016, "BaseToDecimal", "Cannot convert to decimal. Value out of bound."
    Else
        Err.Raise 1019, "BaseToDecimal", "Cannot convert to decimal due to an unknown error."
    End If
End Function

'Convert from any binary string to decimal
'Need the Microsoft VBScript Regular Expressions 5.5 reference to test if string is a binary string
'Base limit is 35 (9 digits (0 is excluded) + 26 letters)
'Parameters :
'- strBin   : the binary string to convert
'Returns : the value (Long) in decimal base (base 10) or -1 if cannot convert
Public Function BinaryToDecimal(ByVal strBin As String) As Long
    Dim X                               As Integer
    Dim reg                             As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^((0|1)*)$"
    
    If (Not reg.test(strBin)) Then
        MsgBox "Cannot convert to decimal value because string isn't a binary number.", vbOKOnly + vbCritical, "Invalid string"
        BinaryToDecimal = -1
        Exit Function
    End If
    
    On Error GoTo cannotConvert
    
    For X = 0 To Len(strBin) - 1
        BinaryToDecimal = CDec(BinaryToDecimal) + val(Mid(strBin, Len(strBin) - X, 1)) * 2 ^ X
    Next X
    
    Exit Function
    
cannotConvert:
    If (Err.number = 13) Then
        Err.Raise 1013, "BinaryToDecimal", "Cannot convert to decimal. Please check that your value is a valid binary string."
    ElseIf (Err.number = 6) Then
        Err.Raise 1016, "BinaryToDecimal", "Cannot convert to decimal. Value out of bound."
    Else
        Err.Raise 1019, "BinaryToDecimal", "Cannot convert to decimal due to an unknown error."
    End If
    
    BinaryToDecimal = -1
End Function

'Convert from any base to any other base
'Using decimal as temporary conversion
'Using baseToDecimal and decimalToBase functions
'Parameters :
'- value    : the value in its source base
'- srcBase  : the source base
'- destBase : the destination base
'Returns : the value (string) in destination base
Public Function ConvertBase(ByVal value As String, ByVal srcBase As Integer, ByVal destBase As Integer) As String
    ConvertBase = DecimalToBase(val(BaseToDecimal(value, srcBase)), destBase)
End Function

'Convert a decimal value to any other base
'Base limit is 35 (9 digits (0 is excluded) + 26 letters)
'Parameters :
'- value    : the value in decimal base
'- destBase : the destination base
'Returns : the value (string) in destination base
Public Function DecimalToBase(ByVal srcValue As Long, ByVal destBase As Integer) As String
    Dim valueRest                       As Long
    Dim toDivide                        As Long
    Dim charRest                        As String
    
    On Error GoTo cannotConvert
    
    toDivide = srcValue
    DecimalToBase = vbNullString
    
    If ((destBase < 2) Or (destBase > 36)) Then
        Err.Raise 1010, "DecimalToBase", "Cannot convert to base over 36 or below 2."
    End If
    
    While (toDivide > 0)
        valueRest = toDivide - Int(toDivide / destBase) * destBase
        charRest = Mid("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", valueRest + 1, 1)
        DecimalToBase = Trim(charRest) & DecimalToBase
        toDivide = Int(toDivide / destBase)
    Wend
    
    Exit Function
    
cannotConvert:
    If (Err.number = 13) Then
        Err.Raise 1013, "DecimalToBase", "Cannot convert to decimal. Please check that your value is a valid number."
    ElseIf (Err.number = 6) Then
        Err.Raise 1016, "DecimalToBase", "Cannot convert to base" & destBase & ". Value out of bound."
    Else
        Err.Raise 1019, "DecimalToBase", "Cannot convert to base" & destBase & " due to an unknown error."
    End If
End Function

'Convert number to binary string
'Parameters :
'- decValue : source number in decimal base
'Returns : a string in binary base or empty string if error occurred
Public Function DecimalToBinary(ByVal decValue As Long) As String
    Dim isNeg                           As Boolean
    
    On Error GoTo cannotConvert
    
    isNeg = (decValue < 0)
    
    If isNeg Then
        decValue = decValue * -1
    End If
    
    While (decValue <> 0)
        DecimalToBinary = Format(decValue - 2 * Int(decValue / 2)) & DecimalToBinary
        decValue = Int(decValue / 2)
    Wend
    
    If isNeg Then
        DecimalToBinary = "-" & DecimalToBinary
    End If
    
    Exit Function
    
cannotConvert:
    If (Err.number = 6) Then
        Err.Raise 1016, "DecimalToBinary", "Cannot convert to binary. Source value out of bound."
    Else
        Err.Raise 1019, "DecimalToBinary", "Cannot convert to binary due to an unknown error."
    End If
    
    DecimalToBinary = vbNullString
End Function
