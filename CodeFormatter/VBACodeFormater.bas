Attribute VB_Name = "VBACodeFormater"
Option Explicit

' Made by Hugues DUMONT
' This project is still in development, therefore, there is only a basis of what the future project will be like

Private Const LINE_BEFORE As String = "For ,While ,If ,ElseIf ,Else ,With ,Select Case ,Case ,Do While ,Do Until ,Public ,Private "
Private Const LINE_AFTER As String = "Next ,End If,End With,Wend ,End Select ,Loop ,End Sub,End Function,End Property,End Enum,End Type"

Public Sub FormatAllCode()
    Dim vbComp                          As VBComponent
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If (vbComp.Name <> "VBACodeFormater") Then
            Call CorrectIf(vbComp.Name)
            Call CorrectInLines(vbComp.Name)
            Call CorrectScope(vbComp.Name)
            Call OrganizeVar(vbComp.Name)
            Call CorrectBlankLines(vbComp.Name)
            Call CorrectIndent(vbComp.Name)
        End If
    Next vbComp
    
    Debug.Print "fini"
End Sub

' Procedure to correct indentation
Public Sub CorrectIndent(ByVal moduleName As String)
    Dim cmod                            As CodeModule
    Dim ll_IndexLine                    As Long
    Dim ls_trimline                     As String
    Dim li_CountTab                     As Integer
    Dim lb_TabAdded                     As Boolean
    Dim lb_FirstCase                    As Boolean
    
    On Error Resume Next
    
    Set cmod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
    
    li_CountTab = 0
    lb_TabAdded = False
    
    For ll_IndexLine = 1 To cmod.CountOfDeclarationLines
        ls_trimline = Trim(cmod.Lines(ll_IndexLine, 1))
        
        If (StartWithList(ls_trimline, Array("Public Enum ", "Private Enum ", "Public Type ", "Private Type "), False, False)) Then
            cmod.ReplaceLine ll_IndexLine, ls_trimline
            li_CountTab = 1
            lb_TabAdded = True
        ElseIf (StartWithList(ls_trimline, Array("End Enum", "End Type"), False, False)) Then
            cmod.ReplaceLine ll_IndexLine, ls_trimline
            li_CountTab = 0
            lb_TabAdded = False
        ElseIf ((StringIsCode(Trim(cmod.Lines(ll_IndexLine - 1, 1)), " _") > 0)) Then
            If (ll_IndexLine > 1) Then
                cmod.ReplaceLine ll_IndexLine, Space(li_CountTab * 4 + IIf(lb_TabAdded, 4, 8)) & ls_trimline
            Else
                cmod.ReplaceLine ll_IndexLine, ls_trimline
            End If
            
            lb_TabAdded = False
        Else
            cmod.ReplaceLine ll_IndexLine, Space(li_CountTab * 4) & ls_trimline
            lb_TabAdded = False
        End If
    Next ll_IndexLine
    
    li_CountTab = 0
    lb_TabAdded = False
    lb_FirstCase = False
    
    For ll_IndexLine = cmod.CountOfDeclarationLines + 1 To cmod.CountOfLines - 1
        ls_trimline = Trim(cmod.Lines(ll_IndexLine, 1))
        
        If (StartWithList(ls_trimline, Array("For ", "While ", "If ", "With ", "Select Case ", "Do While ", "Do Until ", "Public ", "Private "), False, False) Or _
                (ls_trimline = "Do") Or (StartWith(ls_trimline, "Do ", False, False) And (CommentStart(cmod.Lines(ll_IndexLine, 1)) - 1 > 0))) Then
            cmod.ReplaceLine ll_IndexLine, Space(li_CountTab * 4) & ls_trimline
            
            lb_FirstCase = StartWith(ls_trimline, "Select Case ", False, False)
            
            If ((StringIsCode(ls_trimline, " Then") = 0) Or ((StringIsCode(ls_trimline, " Then") > 0) And ((StringIsCode(ls_trimline, " Then  ") > 0) Or (StringIsCode(ls_trimline, " Then '") > 0) Or (Right(ls_trimline, 5) = " Then")))) Then ' Then sans instruction sur la même ligne
                li_CountTab = li_CountTab + 1
                lb_TabAdded = True
            Else
                lb_TabAdded = False
            End If
        ElseIf ((StringIsCode(ls_trimline, "Else") > 0) Or StartWithList(ls_trimline, Array("ElseIf ", "Case "), False, False)) Then
            If (lb_FirstCase) Then
                cmod.ReplaceLine ll_IndexLine, Space(li_CountTab * 4) & ls_trimline
                li_CountTab = li_CountTab + 1
                lb_FirstCase = False
            Else
                cmod.ReplaceLine ll_IndexLine, Space((li_CountTab - 1) * 4) & ls_trimline
            End If
            
            lb_TabAdded = True
        ElseIf (StringIsCode(ls_trimline, "End Select") > 0) Then
            li_CountTab = li_CountTab - 2
            lb_TabAdded = False
            cmod.ReplaceLine ll_IndexLine, Space(li_CountTab * 4) & ls_trimline
        ElseIf (StartWithList(ls_trimline, Array("End Function", "End Sub", "End Property", "End Enum", "End Type", "End With", "End If", "Next ", "Loop "), False, False) Or _
                (ls_trimline = "Next") Or (ls_trimline = "Wend") Or (ls_trimline = "Loop")) Then
            li_CountTab = li_CountTab - 1
            lb_TabAdded = False
            cmod.ReplaceLine ll_IndexLine, Space(li_CountTab * 4) & ls_trimline
        ElseIf (StringIsCode(Trim(cmod.Lines(ll_IndexLine - 1, 1)), " _") > 0) Then
            cmod.ReplaceLine ll_IndexLine, Space(li_CountTab * 4 + IIf(lb_TabAdded, 4, 8)) & ls_trimline
            lb_TabAdded = False
        Else
            cmod.ReplaceLine ll_IndexLine, Space(li_CountTab * 4) & ls_trimline
            lb_TabAdded = False
        End If
        
        If (li_CountTab <= 0) Then
            li_CountTab = 0
        End If
    Next ll_IndexLine
    
    cmod.ReplaceLine cmod.CountOfLines, Trim(cmod.Lines(cmod.CountOfLines, 1))
End Sub

' Procedure to correct blank lines in a component
Public Sub CorrectBlankLines(ByVal moduleName As String)
    Dim cmod                            As CodeModule
    Dim ll_StartDeclareLine             As Long
    Dim li_CommPos                      As Integer
    Dim ls_TrimPrevious                 As String
    Dim ls_trimline                     As String
    Dim ls_CodeLine                     As String
    Dim ll_IndexLine                    As Long
    
    On Error Resume Next
    
    Set cmod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
    
    For ll_IndexLine = 1 To cmod.CountOfDeclarationLines
        ls_CodeLine = cmod.Lines(ll_IndexLine, 1)
        ls_trimline = Trim(ls_CodeLine)
        li_CommPos = CommentStart(ls_CodeLine) - 1
        
        If (StartWith(ls_trimline, "Option ")) Then
            If ((Not StartWith(Trim(cmod.Lines(ll_IndexLine + 1, 1)), "Option ")) And (Trim(cmod.Lines(ll_IndexLine + 1, 1) <> ""))) Then
                cmod.ReplaceLine ll_IndexLine, ls_trimline & Chr(13)
            End If
        ElseIf (StartWith(ls_trimline, "Implements ", False, False)) Then
            If ((Not StartWith(Trim(cmod.Lines(ll_IndexLine + 1, 1)), "Implements ", False, False)) And (Trim(cmod.Lines(ll_IndexLine + 1, 1)) <> "")) Then
                cmod.ReplaceLine ll_IndexLine, ls_trimline & Chr(13)
            End If
            
            If ((Not StartWithList(Trim(cmod.Lines(ll_IndexLine - 1, 1)), Array("'", "Implements "), False, False)) And (Trim(cmod.Lines(ll_IndexLine - 1, 1)) <> "")) Then
                cmod.ReplaceLine ll_IndexLine, Chr(13) & ls_trimline
            End If
        ElseIf (StartWithList(ls_trimline, Array("Public ", "Private "), False, False)) Then
            If (Trim(cmod.Lines(ll_IndexLine + 1, 1)) = "") Then
                While ((Trim(cmod.Lines(ll_IndexLine + 2, 1)) = "") And (ll_IndexLine + 2 < cmod.CountOfLines))
                    cmod.DeleteLines ll_IndexLine + 2, 1
                Wend
                
                If (Not StartWithList(Trim(cmod.Lines(ll_IndexLine + 2, 1)), Array("'", "Public ", "Private ", "Implements "), False, False)) Then
                    cmod.DeleteLines ll_IndexLine + 1, 1
                End If
            End If
        ElseIf (Left(ls_trimline, 1) = "'") Then
            If (Trim(cmod.Lines(ll_IndexLine + 1, 1)) = "") Then
                cmod.ReplaceLine ll_IndexLine, ls_CodeLine & Chr(13)
            End If
            
            If ((Left(Trim(cmod.Lines(ll_IndexLine - 1, 1)), 1) <> "'") And (Trim(cmod.Lines(ll_IndexLine - 1, 1)) <> "")) Then
                cmod.ReplaceLine ll_IndexLine, Chr(13) & ls_CodeLine
            End If
        End If
    Next ll_IndexLine
    
    ll_IndexLine = cmod.CountOfDeclarationLines + 1
    
    While (ll_IndexLine <= cmod.CountOfLines)
        ls_CodeLine = cmod.Lines(ll_IndexLine, 1)
        ls_trimline = Trim(ls_CodeLine)
        
        If (StartWithList(ls_trimline, Array("Public ", "Private "), False, False)) Then
            ll_StartDeclareLine = ll_IndexLine
            
            While (StringIsCode(cmod.Lines(ll_StartDeclareLine, 1), " _") > 0)
                ll_StartDeclareLine = ll_StartDeclareLine + 1
            Wend
            
            ll_StartDeclareLine = ll_StartDeclareLine + 1
            
            While (Trim(cmod.Lines(ll_StartDeclareLine, 1)) = "")
                cmod.DeleteLines ll_StartDeclareLine, 1
            Wend
            
            If ((Left(Trim(cmod.Lines(ll_IndexLine - 1, 1)), 1) <> "'") And (Trim(cmod.Lines(ll_IndexLine - 1, 1)) <> "")) Then
                cmod.ReplaceLine ll_IndexLine, Chr(13) & ls_trimline
                ll_IndexLine = ll_IndexLine + 1
            End If
        ElseIf (StartWithList(ls_trimline, Split(LINE_BEFORE, ","), False, False) Or (ls_trimline = "Do") Or (StartWith(ls_trimline, "Do ", False, False) And (li_CommPos > 0)) Or _
                (StringIsCode(ls_CodeLine, ":") > 0) Or (StringIsCode(ls_trimline, "Else") > 0)) Then
            If (StartWithList(ls_trimline, Array("Case ", "ElseIf "), False, False) Or StringIsCode(ls_trimline, "Else")) Then
                While ((Trim(cmod.Lines(ll_IndexLine - 1, 1)) = "") And (ll_IndexLine > 1))
                    cmod.DeleteLines ll_IndexLine - 1, 1
                    ll_IndexLine = ll_IndexLine - 1
                Wend
            Else
                ls_TrimPrevious = Trim(cmod.Lines(ll_IndexLine - 1, 1))
                
                If ((Not (StartWithList(ls_TrimPrevious, Split(LINE_BEFORE, ","), False, False) Or (ls_TrimPrevious = "Do") Or (StartWith(ls_TrimPrevious, "Do ", False, False) And (li_CommPos > 0)))) And _
                        (ls_TrimPrevious <> "") And (Left(ls_TrimPrevious, 1) <> "'") And (StringIsCode(ls_TrimPrevious, " Then") = 0) And (StringIsCode(ls_TrimPrevious, ":") = 0) And _
                            (StringIsCode(ls_TrimPrevious, "Else") = 0)) Then
                    cmod.ReplaceLine ll_IndexLine, Chr(13) & ls_CodeLine
                    ll_IndexLine = ll_IndexLine + 1
                End If
            End If
            
            While (StringIsCode(cmod.Lines(ll_IndexLine, 1), " _") > 0)
                ll_IndexLine = ll_IndexLine + 1
            Wend
            
            If ((StringIsCode(ls_CodeLine, " Then  ") > 0) Or (StringIsCode(ls_CodeLine, " Then '") > 0) Or (StringIsCode(ls_CodeLine, " Then") = 0)) Then ' then sans instruction sur la même ligne ou pas de then
                While (Trim(cmod.Lines(ll_IndexLine + 1, 1)) = "")
                    cmod.DeleteLines ll_IndexLine + 1, 1
                Wend
            End If
        ElseIf (StringIsCode(ls_CodeLine, " Then") > 0) Then
            If ((StringIsCode(ls_CodeLine, " Then  ") > 0) Or (StringIsCode(ls_CodeLine, " Then '") > 0)) Then ' then sans instruction sur la même ligne
                While (Trim(cmod.Lines(ll_IndexLine + 1, 1)) = "")
                    cmod.DeleteLines ll_IndexLine + 1, 1
                Wend
            End If
        ElseIf (StartWithList(ls_trimline, Split(LINE_AFTER, ","), False, False) Or (ls_trimline = "Next") Or (ls_trimline = "Wend") Or (ls_trimline = "Loop")) Then
            While ((Trim(cmod.Lines(ll_IndexLine - 1, 1)) = "") And (ll_IndexLine > 1))
                cmod.DeleteLines ll_IndexLine - 1, 1
                ll_IndexLine = ll_IndexLine - 1
            Wend
            
            If (Trim(cmod.Lines(ll_IndexLine + 1, 1)) <> "") Then
                cmod.ReplaceLine ll_IndexLine, ls_CodeLine & Chr(13)
            End If
        End If
        
        ll_IndexLine = ll_IndexLine + 1
    Wend
    
    For ll_IndexLine = cmod.CountOfLines To 2 Step -1
        If (Trim(cmod.Lines(ll_IndexLine, 1)) = "") Then
            If (Trim(cmod.Lines(ll_IndexLine - 1, 1)) = "") Then
                cmod.DeleteLines ll_IndexLine - 1, 1
            End If
        End If
    Next ll_IndexLine
    
    While ((Trim(cmod.Lines(ll_IndexLine, 1)) = "") And (cmod.CountOfLines > 0))
        cmod.DeleteLines ll_IndexLine
    Wend
    
    While ((Trim(cmod.Lines(cmod.CountOfLines, 1)) = "") And (cmod.CountOfLines > 0))
        cmod.DeleteLines cmod.CountOfLines
    Wend
End Sub

' Procedure to correct "if" instructions containing colon (:)
Public Sub CorrectIf(ByVal moduleName As String)
    Dim cmod                            As CodeModule
    Dim ll_IndexLine                    As Long
    Dim li_StringPos                    As Integer
    
    Set cmod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
    
    For ll_IndexLine = 1 To cmod.CountOfLines
        li_StringPos = StringIsCode(cmod.Lines(ll_IndexLine, 1), "Then:")
        
        If (li_StringPos > 0) Then
            cmod.ReplaceLine li_StringPos, Left(cmod.Lines(ll_IndexLine, 1), li_StringPos - 1) & "Then" & Mid(cmod.Lines(li_StringPos, 1), li_StringPos + 5)
        End If
        
        li_StringPos = StringIsCode(cmod.Lines(ll_IndexLine, 1), "Else:")
        
        If (li_StringPos > 0) Then
            cmod.ReplaceLine li_StringPos, Left(cmod.Lines(ll_IndexLine, 1), li_StringPos - 1) & "Else" & Mid(cmod.Lines(li_StringPos, 1), li_StringPos + 5)
        End If
    Next ll_IndexLine
End Sub

' Procedure to correct multiple instructions in a single line (separated with colon (:)) into multiple lines (one instruction by line)
Public Sub CorrectInLines(ByVal moduleName As String)
    Dim cmod                            As CodeModule
    Dim ls_Spaces                       As String
    Dim li_CommPos                      As Integer
    Dim li_ColonPos                     As Integer
    Dim ls_CodeLine                     As String
    Dim ll_IndexLine                    As Long
    
    Set cmod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
    
    For ll_IndexLine = 1 To cmod.CountOfLines
        ls_CodeLine = cmod.Lines(ll_IndexLine, 1)
        li_ColonPos = StringIsCode(ls_CodeLine, ":")
        
        If (li_ColonPos > 0) Then ' Multiple instructions in a single line or etiquette
            
            If (StringIsCode(Mid(ls_CodeLine, li_ColonPos, 2), ":=") = 0) Then
                li_CommPos = CommentStart(ls_CodeLine) - 1
                
                If (InStr(RTrim(Left(ls_CodeLine, IIf(li_CommPos > 0, li_CommPos, Len(ls_CodeLine)))), " ") > 0) Then ' If there are spaces before start of comment then there are multiple instructions
                    ls_Spaces = Left(ls_CodeLine, Len(ls_CodeLine) - Len(Trim(ls_CodeLine)))
                    cmod.ReplaceLine ll_IndexLine, Left(ls_CodeLine, li_ColonPos - 1) & Chr(13) & ls_Spaces & Trim(Mid(ls_CodeLine, li_ColonPos + 1))
                End If
            End If
        End If
    Next ll_IndexLine
End Sub

' Procedure to reorganize var in a module
Public Sub OrganizeVar(ByVal moduleName As String)
    Dim cmod                            As CodeModule
    Dim ll_StartDeclareLine             As Long
    Dim lv_VarArr                       As Variant
    Dim ls_trimline                     As String
    Dim ls_CodeLine                     As String
    Dim ll_ProcLine                     As Long
    Dim ll_ProcStart                    As Long
    Dim li_start                        As Integer
    Dim li_DeclarePos                   As Integer
    Dim li_CommPos                      As Integer
    
    Set cmod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
    
    ' Separate multiple declaration to one declaration by line
    For ll_ProcLine = cmod.CountOfDeclarationLines + 1 To cmod.CountOfLines
        ls_CodeLine = cmod.ProcOfLine(ll_ProcLine, vbext_pk_Proc)
        
        On Error Resume Next
        li_start = 0
        
        While (li_start < 4)
            ll_ProcStart = cmod.ProcBodyLine(ls_CodeLine, li_start)
            
            If (Err.Number = 0) Then
                li_start = 5
            Else
                On Error GoTo 0
                On Error Resume Next
                li_start = li_start + 1
            End If
        Wend
        
        If (li_start = 4) Then
            Exit Sub
        End If
        
        ls_CodeLine = cmod.Lines(ll_ProcLine, 1)
        ls_trimline = Trim(ls_CodeLine)
        li_CommPos = CommentStart(ls_trimline) - 1
        
        If (StartWithList(ls_trimline, Array("Private ", "Public "), False, False)) Then
            ll_StartDeclareLine = ll_ProcLine
            
            While (StringIsCode(cmod.Lines(ll_StartDeclareLine, 1), " _") > 0)
                ll_StartDeclareLine = ll_StartDeclareLine + 1
            Wend
            
            ll_StartDeclareLine = ll_StartDeclareLine + 1
            
            While (Left(Trim(cmod.Lines(ll_StartDeclareLine, 1)), 1) = "'")
                ll_StartDeclareLine = ll_StartDeclareLine + 1
            Wend
        End If
        
        If (StartWith(ls_trimline, "Dim ", False, False)) Then ' Line starts with var declaration
            If (StringIsCode(ls_trimline, ",") > 0) And (Not ((StringIsCode(ls_trimline, "(") < StringIsCode(ls_trimline, ",")) And (StringIsCode(ls_trimline, "(") > 0))) Then
                lv_VarArr = Split(Left(ls_trimline, IIf(li_CommPos > 0, li_CommPos, Len(ls_trimline))), ",")
                ls_trimline = Space(4) & lv_VarArr(0)
                
                If (li_CommPos > 0) Then
                    ls_trimline = ls_trimline & Mid(ls_trimline, li_CommPos)
                End If
                
                For li_DeclarePos = 1 To UBound(lv_VarArr)
                    ls_trimline = ls_trimline & Chr(13) & Space(4) & "Dim " & Trim(lv_VarArr(li_DeclarePos)) & IIf(InStr(lv_VarArr(li_DeclarePos), " As ") > 0, "", " As Variant")
                Next li_DeclarePos
                
                If (ll_StartDeclareLine <> ll_ProcLine) Then
                    cmod.DeleteLines ll_ProcLine, 1
                    cmod.ReplaceLine ll_StartDeclareLine, cmod.Lines(ll_StartDeclareLine, 1) & Chr(13) & ls_trimline
                Else
                    cmod.ReplaceLine ll_ProcLine, ls_trimline & Chr(13)
                End If
            End If
        ElseIf (StartWith(ls_trimline, "Const ", False, False)) Then ' Line start with const declaration
            lv_VarArr = Array()
            Call StringCodePositions(ls_CodeLine, ",", lv_VarArr)
            
            If ((Not IsEmpty(lv_VarArr)) And (UBound(lv_VarArr) > 0)) Then
                li_start = 1
                ls_trimline = ""
                
                For li_DeclarePos = 0 To UBound(lv_VarArr)
                    ls_trimline = ls_trimline & Chr(13) & Trim(Mid(ls_CodeLine, li_start, (lv_VarArr(li_DeclarePos) - li_start)))
                    li_start = lv_VarArr(li_DeclarePos) + 1
                Next li_DeclarePos
                
                ls_trimline = Space(4) & Mid(ls_trimline, 2) & Chr(13) & Space(4) & "Const " & Trim(Mid(ls_CodeLine, lv_VarArr(li_DeclarePos - 1) + 1))
            Else
                ls_trimline = Space(4) & ls_trimline
            End If
            
            If (ll_StartDeclareLine <> ll_ProcLine) Then
                cmod.DeleteLines ll_ProcLine, 1
                cmod.ReplaceLine ll_StartDeclareLine, cmod.Lines(ll_StartDeclareLine, 1) & Chr(13) & ls_trimline & Chr(13)
            Else
                cmod.ReplaceLine ll_ProcLine, ls_trimline & Chr(13)
            End If
        End If
    Next ll_ProcLine
    
    ' Align "As" declarations vertically (to 40 characters)
    For ll_ProcLine = cmod.CountOfDeclarationLines + 1 To cmod.CountOfLines
        ls_CodeLine = cmod.Lines(ll_ProcLine, 1)
        
        If (StartWithList(Trim(ls_CodeLine), Array("Dim ", "Const "), False, False)) Then
            li_start = StringIsCode(ls_CodeLine, "As ")
            
            If (li_start > 0) Then
                While (li_start < 41)
                    ls_CodeLine = Left(ls_CodeLine, li_start - 1) & " " & Mid(ls_CodeLine, li_start)
                    li_start = StringIsCode(ls_CodeLine, "As ")
                Wend
                
                cmod.ReplaceLine ll_ProcLine, ls_CodeLine
            End If
        End If
    Next ll_ProcLine
End Sub

' Procedure to explicitly declare scope
Public Sub CorrectScope(ByVal moduleName As String)
    Dim cmod                            As CodeModule
    Dim ll_IndexLine                    As Long
    Dim ls_trimline                     As String
    
    Set cmod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
    
    For ll_IndexLine = 1 To cmod.CountOfLines
        ls_trimline = Trim(cmod.Lines(ll_IndexLine, 1))
        
        If (Trim(cmod.ProcOfLine(ll_IndexLine, vbext_pk_Proc)) = vbNullString) Then
            If (StartWith(ls_trimline, "Dim ")) Then
                cmod.ReplaceLine ll_IndexLine, "Private " & Right(ls_trimline, Len(ls_trimline) - 4)
            ElseIf (StartWith(ls_trimline, "Const ", False, False)) Then
                cmod.ReplaceLine ll_IndexLine, "Private " & ls_trimline
            ElseIf (StartWithList(ls_trimline, Array("Enum ", "Declare ", "Event ", "WithEvents ", "Property ", "Static ", "Type "), False, False)) Then
                cmod.ReplaceLine ll_IndexLine, "Public " & ls_trimline
            ElseIf (StartWithList(ls_trimline, Array("Implements ", "End Enum", "End Type", "Option "), False, False)) Then
                cmod.ReplaceLine ll_IndexLine, ls_trimline
            End If
        ElseIf (StartWithList(ls_trimline, Array("Sub ", "Function ", "Static ", "Property ", "Event ", "WithEvents ", "Enum ", "Type ", "Declare "), False, False)) Then
            If ((StringIsCode(ls_trimline, "Static ") = 0) Or _
                ((StringIsCode(ls_trimline, "Static ") > 0) And StartWithList(Replace(ls_trimline, "Static ", ""), Array("Function ", "Sub ", "Property "), False, False))) Then
                cmod.ReplaceLine ll_IndexLine, "Public " & ls_trimline
            End If
        End If
    Next ll_IndexLine
End Sub

' Procedure to check if string is not between quotation marks and not in comment
Public Function StringIsCode(ByVal ps_CodeLine As String, ByVal ps_String As String) As Integer
    Dim li_pos                          As Long
    Dim lb_Inquote                      As Boolean
    Dim li_CodeEnd                      As Long
    
    lb_Inquote = False
    StringIsCode = 0
    li_CodeEnd = CommentStart(ps_CodeLine)
    
    If (li_CodeEnd = 0) Then
        li_CodeEnd = Len(ps_CodeLine)
    End If
    
    For li_pos = 1 To li_CodeEnd
        If (Mid(ps_CodeLine, li_pos, 1) = Chr(34)) Then
            lb_Inquote = (Not lb_Inquote)
        ElseIf (Mid(ps_CodeLine, li_pos, Len(ps_String)) = ps_String) Then
            If (Not lb_Inquote) Then
                StringIsCode = li_pos
                Exit Function
            End If
        End If
    Next li_pos
End Function

' Procedure to check if string is not between quotation marks and not in comment
Public Sub StringCodePositions(ByVal ps_CodeLine As String, ByVal ps_String As String, ByRef ps_SplitPos As Variant)
    Dim li_pos                          As Long
    Dim lb_Inquote                      As Boolean
    Dim li_CodeEnd                      As Long
    
    lb_Inquote = False
    li_CodeEnd = CommentStart(ps_CodeLine)
    
    If (li_CodeEnd = 0) Then
        li_CodeEnd = Len(ps_CodeLine)
    End If
    
    ps_SplitPos = Array()
    
    For li_pos = 1 To li_CodeEnd
        If (Mid(ps_CodeLine, li_pos, 1) = Chr(34)) Then
            lb_Inquote = (Not lb_Inquote)
        ElseIf (Mid(ps_CodeLine, li_pos, Len(ps_String)) = ps_String) Then
            If (Not lb_Inquote) Then
                Call AddArray(ps_SplitPos, li_pos)
            End If
        End If
    Next li_pos
End Sub

' Procedure to get the position of the beginning of the comment
Public Function CommentStart(ByVal ps_CodeLine As String) As Integer
    Dim li_pos                          As Long
    Dim ls_Char                         As String
    Dim lb_Inquote                      As Boolean
    
    lb_Inquote = False
    CommentStart = 0
    
    For li_pos = 1 To Len(ps_CodeLine)
        ls_Char = Mid(ps_CodeLine, li_pos, 1)
        
        If (ls_Char = Chr(34)) Then
            lb_Inquote = (Not lb_Inquote)
        ElseIf (ls_Char = "'") Then
            If (Not lb_Inquote) Then
                CommentStart = li_pos
                Exit Function
            End If
        End If
    Next li_pos
End Function
