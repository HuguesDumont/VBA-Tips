Attribute VB_Name = "StringInt"
Attribute VB_Description = "Functions for String with Integer type"
Option Explicit

'Function to check if string is integer below zero
'It uses the "isInteger" function implemented before
Public Function IsIntNeg(ByVal value As String) As Boolean
    IsIntNeg = False
    
    If IsInteger(value) Then
        IsIntNeg = (CInt(value) < 0)
    End If
End Function

'Function to check if string is integer above or equal zero
'It uses the "isInteger" function implemented before
Public Function IsIntNegOrZero(ByVal value As String) As Boolean
    IsIntNegOrZero = False
    
    If IsInteger(value) Then
        IsIntNegOrZero = (CInt(value) <= 0)
    End If
End Function

'Function to check if string is an integer over zero
'It uses the "isInteger" function implemented before
Public Function IsIntPos(ByVal value As String) As Boolean
    IsIntPos = False
    
    If IsInteger(value) Then
        IsIntPos = (CInt(value) > 0)
    End If
End Function

'Function to check if string is integer above or equal zero
'It uses the "isInteger" function implemented before
Public Function IsIntPosOrZero(ByVal value As String) As Boolean
    IsIntPosOrZero = False
    
    If IsInteger(value) Then
        IsIntPosOrZero = (CInt(value) >= 0)
    End If
End Function

'Function to check if string is an integer
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function IsInteger(ByVal value As String) As Boolean
    Const INT_MIN                       As Integer = -32768
    Const INT_MAX                       As Integer = 32767
    
    Dim reg                             As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^(-)?(\d)+$"
    IsInteger = False
    
    If reg.test(value) Then
        On Error GoTo capacityOverflow
        IsInteger = ((CInt(value) >= INT_MIN) And (CInt(value) <= INT_MAX))
    End If
    
    Set reg = Nothing
    Exit Function
    
capacityOverflow:
    MsgBox "Value is integer but over 32 767 or lower than -32 768" & Chr(13) & "Can't be converted to the integer type in vba (might be able to convert to long type)!", vbOKOnly + vbCritical, "Capacity overflow !"
    Set reg = Nothing
End Function

'Function to check if string is zero (integer 0)
'It uses the "isInteger" function implemented before
Public Function IsZero(ByVal value As String) As Boolean
    IsZero = False
    
    If IsInteger(value) Then
        IsZero = (CInt(value) = 0)
    End If
End Function
