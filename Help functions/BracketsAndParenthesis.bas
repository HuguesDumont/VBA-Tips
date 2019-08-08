Attribute VB_Name = "BracketsAndParenthesis"
Option Explicit

'Check string for balanced brackets
'Parameters :
'- str      : string to check for balanced brackets
'Returns : true if string has balanced brackets, else false
Public Function CheckBrackets(ByVal str As String) As Boolean
    Dim Depth As Integer
    Dim i As Long
    Dim ch As String * 1

    Depth = 0
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        If ch = "[" Then
            Depth = Depth + 1
        End If

        If ch = "]" Then
            If Depth = 0 Then
                CheckBrackets = False
                Exit Function
            Else
                Depth = Depth - 1
            End If
        End If
    Next i
    CheckBrackets = (Depth = 0)
End Function

'Check string for balanced parenthesis
'Parameters :
'- str      : string to check for balanced parenthesis
'Returns : true if string has balanced parenthesis, else false
Public Function CheckParenthesis(ByVal str As String) As Boolean
    Dim Depth As Integer
    Dim i As Long
    Dim ch As String * 1

    Depth = 0
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        If ch = "(" Then
            Depth = Depth + 1
        End If

        If ch = ")" Then
            If Depth = 0 Then
                CheckBrackets = False
                Exit Function
            Else
                Depth = Depth - 1
            End If
        End If
    Next i
    CheckBrackets = (Depth = 0)
End Function
