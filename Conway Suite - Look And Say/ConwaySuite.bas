Attribute VB_Name = "ConwaySuite"
Attribute VB_Description = "Implementation of the conway suite (look-and-say suite)"
Option Explicit

'Function to display the nth result of the conway suite.
'If n < 1 then returns -1
Public Function lookAndSay(ByVal n As Long) As String
Attribute lookAndSay.VB_Description = "Function to display the nth result of the conway suite.\r\nIf n < 1 then returns -1"
    Dim i As Long, j As Long, x As Long
    Dim str As String, res As String, char As String
    
    If (n < 1) Then
        str = "-1"
    ElseIf (n = 1) Then
        str = "1"
    Else
        str = "1"
        res = ""
        j = 2
        
        While (j <= n + 1)
            char = Mid(str, 1, 1)
            x = 1
            For i = 1 To Len(str)
                If (i + 1 > Len(str)) Then
                    res = res & CStr(x) & char
                    Exit For
                ElseIf (char = Mid(str, i + 1, 1)) Then
                    x = x + 1
                Else
                    res = res & CStr(x) & char
                    x = 1
                    char = Mid(str, i + 1, 1)
                End If
            Next i
            
            j = j + 1
            str = res
            res = ""
        Wend
    End If
    
    lookAndSay = str
End Function
