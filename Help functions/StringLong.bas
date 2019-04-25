Attribute VB_Name = "StringLong"
Attribute VB_Description = " Functions for String with Long type\n"
Option Explicit

'Function to check if string is a long
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function isLong(value As String) As Boolean
Attribute isLong.VB_Description = "Function to check if string is a long\r\nNeed to activate the reference ""Microsoft VBScrpt Regular Expressions 5.5"""
    Dim reg As New VBScript_RegExp_55.RegExp
    Const LONG_MIN As Long = -2147483648#
    Const LONG_MAX As Long = 2147483647

    reg.Pattern = "^(-)?(\d)+$"
    isLong = False
    If reg.test(value) Then
        On Error GoTo capacityOverflow
        If ((CLng(value) >= LONG_MIN) And (CLng(value) <= LONG_MAX)) Then
            isLong = True
        End If
    End If
    Set reg = Nothing
    Exit Function
capacityOverflow:
    MsgBox "Value is integer but over 2 147 483 647 or lower than –2 147 483 648" & Chr(13) & _
        "Can't be converted to the long type in vba!", _
        vbOKOnly + vbCritical, "Capacity overflow !"
    Set reg = Nothing
End Function

'Function to check if string is a long over zero
'It uses the "isLong" function implemented before
Public Function isLongPos(value As String) As Boolean
Attribute isLongPos.VB_Description = "Function to check if string is a long over zero\r\nIt uses the ""isLong"" function implemented before"
    isLongPos = False
    If isLong(value) Then
        If CLng(value) > 0 Then
            isLongPos = True
        End If
    End If
End Function

'Function to check if string is long below zero
'It uses the "isLong" function implemented before
Public Function isLongNeg(value As String) As Boolean
Attribute isLongNeg.VB_Description = "Function to check if string is long below zero\r\nIt uses the ""isLong"" function implemented before"
    isLongNeg = False
    If isLong(value) Then
        If CLng(value) < 0 Then
            isLongNeg = True
        End If
    End If
End Function

'Function to check if string is zero (integer 0)
'It uses the "isLong" function implemented before
Public Function isZero(value As String) As Boolean
Attribute isZero.VB_Description = "Function to check if string is zero (integer 0)\r\nIt uses the ""isLong"" function implemented before"
    isZero = False
    If isLong(value) Then
        If CLng(value) = 0 Then
            isZero = True
        End If
    End If
End Function

'Function to check if string is long above or equal zero
'It uses the "isLong" function implement before
Public Function isLongPosOrZero(value As String) As Boolean
Attribute isLongPosOrZero.VB_Description = "Function to check if string is long above or equal zero\r\nIt uses the ""isLong"" function implement before"
    isLongPosOrZero = False
    If isLong(value) Then
        If CLng(value) >= 0 Then
            isLongPosOrZero = True
        End If
    End If
End Function

'Function to check if string is long above or equal zero
'It uses the "isLong" function implement before
Public Function isLongNegOrZero(value As String) As Boolean
Attribute isLongNegOrZero.VB_Description = "Function to check if string is long above or equal zero\r\nIt uses the ""isLong"" function implement before"
    isLongNegOrZero = False
    If isLong(value) Then
        If CLng(value) <= 0 Then
            isLongNegOrZero = True
        End If
    End If
End Function
