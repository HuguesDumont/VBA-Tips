Attribute VB_Name = "StringLong"
Attribute VB_Description = " Functions for String with Long type\n"
Option Explicit

'Function to check if string is a long
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function IsLong(ByVal value As String) As Boolean
Attribute IsLong.VB_Description = "Function to check if string is a long\r\nNeed to activate the reference ""Microsoft VBScrpt Regular Expressions 5.5"""
    Dim reg As New VBScript_RegExp_55.RegExp
    Const LONG_MIN As Long = -2147483648#
    Const LONG_MAX As Long = 2147483647

    reg.Pattern = "^(-)?(\d)+$"
    IsLong = False
    If reg.test(value) Then
        On Error GoTo capacityOverflow
        If ((CLng(value) >= LONG_MIN) And (CLng(value) <= LONG_MAX)) Then
            IsLong = True
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
Public Function IsLongPos(ByVal value As String) As Boolean
Attribute IsLongPos.VB_Description = "Function to check if string is a long over zero\r\nIt uses the ""isLong"" function implemented before"
    IsLongPos = False
    If IsLong(value) Then
        If CLng(value) > 0 Then
            IsLongPos = True
        End If
    End If
End Function

'Function to check if string is long below zero
'It uses the "isLong" function implemented before
Public Function IsLongNeg(ByVal value As String) As Boolean
Attribute IsLongNeg.VB_Description = "Function to check if string is long below zero\r\nIt uses the ""isLong"" function implemented before"
    IsLongNeg = False
    If IsLong(value) Then
        If CLng(value) < 0 Then
            IsLongNeg = True
        End If
    End If
End Function

'Function to check if string is zero (integer 0)
'It uses the "isLong" function implemented before
Public Function IsZero(ByVal value As String) As Boolean
Attribute IsZero.VB_Description = "Function to check if string is zero (integer 0)\r\nIt uses the ""isLong"" function implemented before"
    IsZero = False
    If IsLong(value) Then
        If CLng(value) = 0 Then
            IsZero = True
        End If
    End If
End Function

'Function to check if string is long above or equal zero
'It uses the "isLong" function implemented before
Public Function IsLongPosOrZero(ByVal value As String) As Boolean
Attribute IsLongPosOrZero.VB_Description = "Function to check if string is long above or equal zero\r\nIt uses the ""isLong"" function implement before"
    IsLongPosOrZero = False
    If IsLong(value) Then
        If CLng(value) >= 0 Then
            IsLongPosOrZero = True
        End If
    End If
End Function

'Function to check if string is long above or equal zero
'It uses the "isLong" function implemented before
Public Function IsLongNegOrZero(ByVal value As String) As Boolean
Attribute IsLongNegOrZero.VB_Description = "Function to check if string is long above or equal zero\r\nIt uses the ""isLong"" function implement before"
    IsLongNegOrZero = False
    If IsLong(value) Then
        If CLng(value) <= 0 Then
            IsLongNegOrZero = True
        End If
    End If
End Function
