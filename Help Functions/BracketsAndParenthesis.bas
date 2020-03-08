Attribute VB_Name = "BracketsAndParenthesis"
Option Explicit

'Check string for balanced brackets
'Parameters :
'- strVal      : string to check for balanced brackets
'Returns : true if string has balanced brackets, else false
Public Function CheckBrackets(ByVal strVal As String) As Boolean
    Dim Depth                           As Integer
    Dim i                               As Long
    Dim Ch                              As String * 1
    
    Depth = 0
    
    For i = 1 To Len(strVal)
        Ch = Mid(strVal, i, 1)
        
        If (Ch = "[") Then
            Depth = Depth + 1
        End If
        
        If (Ch = "]") Then
            If (Depth = 0) Then
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
'- strVal      : string to check for balanced parenthesis
'Returns : true if string has balanced parenthesis, else false
Public Function CheckParenthesis(ByVal strVal As String) As Boolean
    Dim Depth                           As Integer
    Dim i                               As Long
    Dim Ch                              As String * 1
    
    Depth = 0
    
    For i = 1 To Len(strVal)
        Ch = Mid(strVal, i, 1)
        
        If (Ch = "(") Then
            Depth = Depth + 1
        End If
        
        If (Ch = ")") Then
            If (Depth = 0) Then
                CheckParenthesis = False
                Exit Function
            Else
                Depth = Depth - 1
            End If
        End If
    Next i
    
    CheckParenthesis = (Depth = 0)
End Function
