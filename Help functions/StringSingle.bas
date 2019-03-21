Attribute VB_Name = "StringSingle"
Option Explicit

'Function to check if string is of single type
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function isSingle(value As String) As Boolean
Attribute isSingle.VB_Description = "Test description"
    Dim reg As New VBScript_RegExp_55.RegExp
    Const SINGLE_MIN As Single = -3.402823E+38
    Const SINGLE_MAX As Single = 1.401298E+45
    
    reg.Pattern = "^(-)?(\d)+(((\.)|(,))(\d)+)?$"
    isSingle = False
    If reg.test(value) Then
        On Error GoTo capacityOverflow
        If ((CSng(value) >= SINGLE_MIN) And (CSng(value) <= SINGLE_MAX)) Then
            isSingle = True
        End If
    End If
    Set reg = Nothing
    Exit Function
capacityOverflow:
    MsgBox "Value is single but over 1.401298E+45 or lower than -3.402823E+38" & Chr(13) & _
        "Can't be converted to the single type in vba!", _
        vbOKOnly + vbCritical, "Capacity overflow !"
    Set reg = Nothing
End Function

'Function to check if string is a Single over zero
'It uses the "isSingle" function implemented before
Public Function isSinglePos(value As String) As Boolean
    isSinglePos = False
    If isSingle(value) Then
        If CLng(value) > 0 Then
            isSinglePos = True
        End If
    End If
End Function

'Function to check if string is Single below zero
'It uses the "isSingle" function implemented before
Public Function isSingleNeg(value As String) As Boolean
    isSingleNeg = False
    If isSingle(value) Then
        If CLng(value) < 0 Then
            isSingleNeg = True
        End If
    End If
End Function

'Function to check if string is zero (integer 0)
'It uses the "isSingle" function implemented before
Public Function isZero(value As String) As Boolean
    isZero = False
    If isSingle(value) Then
        If CLng(value) = 0 Then
            isZero = True
        End If
    End If
End Function

'Function to check if string is Single above or equal zero
'It uses the "isSingle" function implement before
Public Function isSinglePosOrZero(value As String) As Boolean
    isSinglePosOrZero = False
    If isSingle(value) Then
        If CLng(value) >= 0 Then
            isSinglePosOrZero = True
        End If
    End If
End Function

'Function to check if string is Single above or equal zero
'It uses the "isSingle" function implement before
Public Function isSingleNegOrZero(value As String) As Boolean
    isSingleNegOrZero = False
    If isSingle(value) Then
        If CLng(value) <= 0 Then
            isSingleNegOrZero = True
        End If
    End If
End Function
