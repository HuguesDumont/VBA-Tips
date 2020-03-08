Attribute VB_Name = "StringSingle"
Attribute VB_Description = "Functions for String with Single type\n"
Option Explicit

'Function to check if string is of single type
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function IsSingle(ByVal value As String) As Boolean
    Const SINGLE_MIN                    As Single = -3.402823E+38
    Const SINGLE_MAX                    As Single = 3.4028235E+38
    
    Dim reg                             As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^(-)?(\d)+(((\.)|(,))(\d)+)?$"
    IsSingle = False
    
    If reg.test(value) Then
        On Error GoTo capacityOverflow
        IsSingle = ((CSng(value) >= SINGLE_MIN) And (CSng(value) <= SINGLE_MAX))
    End If
    
    Set reg = Nothing
    Exit Function
    
capacityOverflow:
    MsgBox "Value is single but over 1.401298E+45 or lower than -3.402823E+38" & Chr(13) & "Can't be converted to the single type in vba!", vbOKOnly + vbCritical, "Capacity overflow !"
    Set reg = Nothing
End Function

'Function to check if string is Single below zero
'It uses the "isSingle" function implemented before
Public Function IsSingleNeg(ByVal value As String) As Boolean
    IsSingleNeg = False
    
    If IsSingle(value) Then
        IsSingleNeg = (CLng(value) < 0)
    End If
End Function

'Function to check if string is Single above or equal zero
'It uses the "isSingle" function implemented before
Public Function IsSingleNegOrZero(ByVal value As String) As Boolean
    IsSingleNegOrZero = False
    
    If IsSingle(value) Then
        IsSingleNegOrZero = (CLng(value) <= 0)
    End If
End Function

'Function to check if string is a Single over zero
'It uses the "isSingle" function implemented before
Public Function IsSinglePos(ByVal value As String) As Boolean
    IsSinglePos = False
    
    If IsSingle(value) Then
        IsSinglePos = (CLng(value) > 0)
    End If
End Function

'Function to check if string is Single above or equal zero
'It uses the "isSingle" function implemented before
Public Function IsSinglePosOrZero(ByVal value As String) As Boolean
    IsSinglePosOrZero = False
    
    If IsSingle(value) Then
        IsSinglePosOrZero = (CLng(value) >= 0)
    End If
End Function

'Function to check if string is zero (integer 0)
'It uses the "isSingle" function implemented before
Public Function IsZero(ByVal value As String) As Boolean
    IsZero = False
    
    If IsSingle(value) Then
        IsZero = (CLng(value) = 0)
    End If
End Function
