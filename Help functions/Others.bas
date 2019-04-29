Attribute VB_Name = "Others"
Option Explicit

'Function to check string for balanced brackets
Public Function checkBrackets(str As String) As Boolean
    Dim i As Long
    Dim Depth As Integer
    Dim ch As String * 1
     
    Depth = 0
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        If ch = "[" Then
            Depth = Depth + 1
        End If
        
        If ch = "]" Then
            If Depth = 0 Then
                checkBrackets = False
                Exit Function
            Else
                Depth = Depth - 1
            End If
        End If
    Next i
    checkBrackets = (Depth = 0)
End Function