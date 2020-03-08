Attribute VB_Name = "StringLong"
Attribute VB_Description = " Functions for String with Long type\n"
Option Explicit

'Function to check if string is a long
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function IsLong(ByVal value As String) As Boolean
    Const LONG_MIN                      As Long = -2147483648#
    Const LONG_MAX                      As Long = 2147483647
    
    Dim reg                             As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^(-)?(\d)+$"
    IsLong = False
    
    If reg.test(value) Then
        On Error GoTo capacityOverflow
        IsLong = ((CLng(value) >= LONG_MIN) And (CLng(value) <= LONG_MAX))
    End If
    
    Set reg = Nothing
    Exit Function
    
capacityOverflow:
    MsgBox "Value is integer but over 2 147 483 647 or lower than –2 147 483 648" & Chr(13) & "Can't be converted to the long type in vba!", vbOKOnly + vbCritical, "Capacity overflow !"
    Set reg = Nothing
End Function

'Function to check if string is long below zero
'It uses the "isLong" function implemented before
Public Function IsLongNeg(ByVal value As String) As Boolean
    IsLongNeg = False
    
    If IsLong(value) Then
        IsLongNeg = (CLng(value) < 0)
    End If
End Function

'Function to check if string is long above or equal zero
'It uses the "isLong" function implemented before
Public Function IsLongNegOrZero(ByVal value As String) As Boolean
    IsLongNegOrZero = False
    
    If IsLong(value) Then
        IsLongNegOrZero = (CLng(value) <= 0)
    End If
End Function

'Function to check if string is a long over zero
'It uses the "isLong" function implemented before
Public Function IsLongPos(ByVal value As String) As Boolean
    IsLongPos = False
    
    If IsLong(value) Then
        IsLongPos = (CLng(value) > 0)
    End If
End Function

'Function to check if string is long above or equal zero
'It uses the "isLong" function implemented before
Public Function IsLongPosOrZero(ByVal value As String) As Boolean
    IsLongPosOrZero = False
    
    If IsLong(value) Then
        IsLongPosOrZero = (CLng(value) >= 0)
    End If
End Function

'Function to check if string is zero (integer 0)
'It uses the "isLong" function implemented before
Public Function IsZero(ByVal value As String) As Boolean
    IsZero = False
    
    If IsLong(value) Then
        IsZero = (CLng(value) = 0)
    End If
End Function
