Attribute VB_Name = "CodeFormat"
' To use procedures defined in this module, you need to activate the references :
' - Microsoft Visual Basic For Applications Extensibility 5.3
' - Microsoft VBScript Regular Expressions 5.0
Option Explicit

Private tabSpaces As String 'String representing the tabulations spaces

' Add an element to array
' Parameters :
' - arr      : the array where to add the element
' - val      : the element to add
Private Sub AddArray(ByRef arr As Variant, ByVal val As Variant)
    ReDim Preserve arr(0 To UBound(arr) + 1) As Variant
    arr(UBound(arr)) = val
End Sub

' Add x tabulations at beginning of string
' Parameters :
' - str      : the string to add tabs
' - x        : the number of tabulations to add
' Returns : string
Private Function AddTabs(ByVal str As String, ByVal x As Integer) As String
    If (tabSpaces = "") Then
        tabSpaces = "    "
    End If
    While (x > 0)
        str = tabSpaces & str
        x = x - 1
    Wend
    AddTabs = str
End Function

' Correct blank lines in a module
' Parameters :
' - moduleName : the module name
Public Sub BlankLineCorrection(ByVal moduleName As String)
    Dim codeMod As CodeModule
    Dim i As Long
    Dim prevLine As String, curLine As String, nextLine As String
    Dim prevArr As Variant, nextArr As Variant

    prevArr = Array("Select ", "Case ", "Do ", "With ", "Public ", "Private ", "Sub ", "Function ", "Friend ", "If ", "For ", "While ", "Else", "ElseIf ")
    nextArr = Array("Else", "ElseIf ", "End", "Exit", "Wend", "Next", "Loop", "Case")
    Set codeMod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule

    With codeMod
        For i = .CountOfLines - 1 To 2 Step -1
            prevLine = Trim(.Lines(i - 1, 1))
            curLine = Trim(.Lines(i, 1))
            nextLine = Trim(.Lines(i + 1, 1))

            If (Len(curLine) = 0) And ((Len(prevLine) = 0 Or StartWithList(prevLine, prevArr) Or prevLine = "Do" Or IsTabLabel(prevLine) Or _
                        EndWith(prevLine, " Then") Or EndWith(prevLine, " _")) Or (StartWithList(nextLine, nextArr)) Or _
                            (IsDim(prevLine) And IsDim(nextLine))) Then
                .DeleteLines i, 1
            ElseIf (IsDim(curLine) And (Not StartWith(curLine, "ReDim ")) And (Not IsDim(nextLine)) And Len(nextLine) <> 0) Then
                .InsertLines i + 1, ""
            ElseIf (prevLine = "Option explicit" And Len(curLine) <> 0) Then
                .InsertLines i, ""
            End If
        Next i
        While ((Len(Trim(.Lines(1, 1))) = 0) And (.CountOfLines <> 0))
            .DeleteLines 1, 1
        Wend
    End With
End Sub

' Check if a module name follow convention
' Parameters :
' - moduleName : the module name
' Returns : string containing description of all problems in the module name format
Public Function CheckModuleName(ByVal moduleName As String) As String
    Dim regModuleName As New VBScript_RegExp_55.RegExp

    regModuleName.Pattern = "^([A-Z][a-zA-Z]([\w])*)$"

    If (Len(moduleName) > 30) Then
        CheckModuleName = CheckModuleName & "- The name of module " & moduleName & " is too long (<30 characters) : " & Len(moduleName) & "." & Chr(13)
    End If
    If (Not regModuleName.test(moduleName)) Then
        CheckModuleName = CheckModuleName & "- The name of module " & moduleName & " doesn't comply with convention ^([A-Z][a-zA-Z]([a-ZA-Z0-9]){1,28})$" & Chr(13)
    End If
End Function

' Check procedure length
' Parameters :
' - moduleName : the module name
' Returns : string with all procedures too long
Public Function CheckProcLength(ByVal moduleName As String) As String
    Dim codeMod As CodeModule
    Dim i As Long, count As Long
    Dim procedureName As String, trimLine As String

    Set codeMod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule

    With codeMod
        For i = .CountOfDeclarationLines + 1 To .CountOfLines - 1
            count = 1
            If (IsProcedure(Trim(.Lines(i, 1)))) Then
                procedureName = GetProcedureName(Trim(.Lines(i, 1)))
                trimLine = Trim(.Lines(i + 1, 1))
                While (trimLine <> "End Sub" And trimLine <> "End Function")
                    If (Not (trimLine = "" Or StartWith(trimLine, "'") Or IsDim(trimLine))) Then
                        count = count + 1
                    End If
                    i = i + 1
                    trimLine = Trim(.Lines(i + 1, 1))
                Wend

                If (count > 30) Then
                    CheckProcLength = CheckProcLength & "- Procedure " & procedureName & " in module " & moduleName & " is too long (>30 lines) : " & count & "." & Chr(13)
                End If
            End If
        Next i
    End With
End Function

' Check for non used procedures and unspecified scope procedures
' If procedure is private then check in module only, else check in project
' Parameters :
' - moduleName : the module name
' Returns : all problems of procedures names in module (name format; unused procedure ; missing scope)
Public Function CheckProcedureNames(ByVal moduleName As String) As String
    Dim found As Boolean
    Dim codeMod As CodeModule, modulecode As CodeModule
    Dim i As Long, j As Long
    Dim regProcName As New VBScript_RegExp_55.RegExp
    Dim trimLine As String, procedureName As String
    Dim projComp As VBComponent

    Set codeMod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
    
    regProcName.Pattern = "^([A-Z][a-zA-Z]([\w])*)$"
    With codeMod
    For i = .CountOfDeclarationLines To .CountOfLines
        trimLine = Trim(.Lines(i, 1))
        If (IsProcedure(trimLine)) Then
            found = False
            procedureName = GetProcedureName(trimLine)
            If (Len(procedureName) > 30) Then
                    CheckProcedureNames = CheckProcedureNames & "- The name of procedure " & procedureName & " in module " & moduleName & " is too long (<30 characters) : " & Len(procedureName) & "." & Chr(13)
                End If
                If (Not regProcName.test(procedureName)) Then
                    CheckProcedureNames = CheckProcedureNames & "- The name of procedure " & procedureName & " in module " & moduleName & " doesn't comply with convention " & _
                            "^([A-Z][a-zA-Z]([a-ZA-Z0-9]){1,28})$" & Chr(13)
                End If
                If (StartWith(trimLine, "Private ")) Then
                    found = UsedInModule(.Parent, procedureName, True)
                ElseIf ((StartWithList(trimLine, Array("Public ", "Friend "))) Or (StartWithList(trimLine, Array("Sub ", "Function ")))) Then
                    If ((StartWithList(trimLine, Array("Sub ", "Function "))) And (Mid(Split(trimLine, "(")(1), 1, 1) <> ")")) Then
                        CheckProcedureNames = CheckProcedureNames & "- Scope of procedure " & procedureName & " line " & i & " in module " & moduleName & " is not specified." & Chr(13)
                    End If

                    For Each projComp In ThisWorkbook.VBProject.VBComponents
                        If (UsedInModule(projComp, procedureName, False, moduleName)) Then
                            found = True
                            Exit For
                        End If
                    Next projComp
                End If
                If (Not found) Then
                    CheckProcedureNames = CheckProcedureNames & IIf(StartWith(trimLine, "Private "), "- Private ", "- Public ")
                    CheckProcedureNames = CheckProcedureNames & IIf(InStr(trimLine, " Function ") > 0 Or StartWith(trimLine, "Function "), "Function ", "Sub ")
                    CheckProcedureNames = CheckProcedureNames & procedureName & " in module " & moduleName & " is not used." & Chr(13)
                End If
            End If
        Next i
    End With
End Function

' Function to correct comments in a module
Public Sub CommentCorrection(ByVal moduleName As String)
    Dim codeMod As CodeModule
    Dim count As Integer
    Dim i As Long
    Dim codeLine As String

    Set codeMod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule

    With codeMod
        For i = 1 To .CountOfLines
            codeLine = .Lines(i, 1)
            If (StartWith(.Lines(i, 1), "'", True)) Then
                If (Len(codeLine) = 1) Then
                    .DeleteLines i
                    i = i - 1
                Else
                    .ReplaceLine i, FormatComment(codeLine)
                End If
            End If
        Next i
    End With
End Sub

' Function to copy all procedures within a module (including comments) in an array
Public Function CopyModuleProc(ByVal moduleName As String) As Variant
    Dim i As Long
    Dim tmpArray As Variant

    tmpArray = ProcNamesModule(moduleName)
    For i = 0 To UBound(tmpArray)
        tmpArray(i) = CopyProcedure(tmpArray(i), moduleName)
    Next i
    CopyModuleProc = tmpArray
End Function

' Function to get all procedure text (including comments)
Public Function CopyProcedure(ByVal procName As String, ByVal moduleName As String) As String
    With ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
        CopyProcedure = .Lines(.ProcStartLine(procName, vbext_pk_Proc), .ProcCountLines(procName, vbext_pk_Proc))
    End With
End Function

' Function to correct the indentation of a specific line.
' The count paramater defines the number of indentation on the previous line
Private Function CorrectLineIndent(ByVal codeLine As String, ByRef count As Integer) As String
    codeLine = Trim(codeLine)

    If (codeLine = "") Then
        CorrectLineIndent = ""
    ElseIf (StartWith(codeLine, "Dim ")) Then
        count = 1
        CorrectLineIndent = AddTabs(codeLine, count)
    ElseIf (StartWithList(codeLine, Array("Public ", "Private ", "Sub ", "Function ", "Friend "))) Then
        count = 0
        CorrectLineIndent = AddTabs(codeLine, count)
    ElseIf (codeLine = "End Sub" Or codeLine = "End Function") Then
        CorrectLineIndent = codeLine
        count = count + 1
    ElseIf (IsTabLabel(codeLine)) Then
        CorrectLineIndent = codeLine
        If (count = 0) Then
            count = 1
        End If
    ElseIf ((EndWith(codeLine, " _") Or EndWith(codeLine, " Then")) And (Not (StartWithList(codeLine, Array("If ", "ElseIf "))))) Then
        count = count + 1
        CorrectLineIndent = AddTabs(codeLine, count)
        count = count - 1
    ElseIf ((StartWith(codeLine, "If ") And Not OneLineIf(codeLine)) Or (StartWith(codeLine, "For ") And Not OneLineFor(codeLine)) Or _
                StartWith(codeLine, "While ") Or StartWith(codeLine, "With ") Or codeLine = "Do" Or StartWith(codeLine, "Do ")) Then
        count = count - 1
        CorrectLineIndent = AddTabs(codeLine, count)
    ElseIf (codeLine = "Else" Or StartWith(codeLine, "ElseIf ") Or StartWith(codeLine, "Case ")) Then
        count = count - 1
        CorrectLineIndent = AddTabs(codeLine, count)
        count = count + 1
    ElseIf (StartWith(codeLine, "Select ")) Then
        count = count - 2
        CorrectLineIndent = AddTabs(codeLine, count)
    ElseIf (codeLine = "End Select") Then
        CorrectLineIndent = AddTabs(codeLine, count)
        count = count + 2
    ElseIf ((StartWith(codeLine, "Next ") Or codeLine = "Next") Or StartWithList(codeLine, Array("End ", "Loop Until")) Or codeLine = "End" Or codeLine = "Wend" Or codeLine = "Loop") Then
        CorrectLineIndent = AddTabs(codeLine, count)
        count = count + 1
    Else
        CorrectLineIndent = AddTabs(codeLine, count)
    End If
End Function

' Function to check for Next with var missing
Public Sub CorrectNext(ByVal moduleName As String)
    Dim codeMod As CodeModule
    Dim i As Long
    Dim forStack As New Stack
    Dim codeLine As String, trimLine As String, spaces() As String

    Set codeMod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule

    With codeMod
        For i = 1 To .CountOfLines
            codeLine = .Lines(i, 1)
            trimLine = Trim(codeLine)
            If (StartWith(trimLine, "For ") And Not OneLineFor(trimLine)) Then
                spaces = Split(trimLine, " ")
                If (StartWith(trimLine, "For Each ")) Then
                    forStack.Push spaces(2)
                Else
                    forStack.Push spaces(1)
                End If
            ElseIf (StartWith(trimLine, "Next")) Then
                spaces = Split(codeLine, "Next")
                .ReplaceLine i, (spaces(0) & "Next " & forStack.Pop())
            End If
        Next i
    End With
End Sub

' Function to get number of occurences of string findStr in string str
Private Function CountOccurences(ByVal str As String, ByVal findStr As String, Optional ByVal withCase As Boolean = True) As Long
    Dim i As Long

    If (Not withCase) Then
        str = UCase(str)
        findStr = UCase(findStr)
    End If

    For i = 1 To Len(str)
        If (Mid(str, i, Len(findStr)) = findStr) Then CountOccurences = CountOccurences + 1
    Next i
End Function

' Function to count tabulations based on defined tab spaces
Private Function CountTabulations(ByVal codeLine As String) As Integer
    Dim i As Long

    If (tabSpaces = "") Then
        tabSpaces = "    "
    End If

    i = 1
    While (Mid(codeLine, i, 1) = " " And i <= Len(codeLine))
        If (i Mod Len(tabSpaces) = 0) Then
            CountTabulations = CountTabulations + 1
        End If
        i = i + 1
    Wend
End Function

' Function to check if string ends with other string (str ends with ending) (no trim)
Private Function EndWith(ByVal str As String, ByVal ending As String, Optional ByVal withCase As Boolean = True) As Boolean
    If (Len(ending) > Len(str)) Then
        EndWith = False
    ElseIf (withCase) Then
        EndWith = (Mid(str, Len(str) - Len(ending) + 1) = ending)
    Else
        EndWith = (Mid(UCase(str), Len(str) - Len(ending)) = UCase(ending))
    End If
End Function

' Function to format a comment correctly
Private Function FormatComment(ByVal str As String) As String
    FormatComment = "'" & Trim(Mid(str, 2))
    If (Len(FormatComment) > 2) Then
        FormatComment = "' " & UCase(Mid(FormatComment, 2, 1)) & Mid(FormatComment, 3)
    End If
End Function

' Function to format the code inside a module :
' - Check if module name is conventional (^([A-Z][a-zA-Z]([a-ZA-Z0-9]){1,28})$)
' - Sorting all procedures by name in lexicographical order (0-9->A-Z->a-z)
' - Correcting the indentation
' - Correcting comment case (first letter has to be upper case)
' - Correcting blank lines
' - Reorganizing var by types
' - Correcting "Next" instruction without var
' - Checking for non-used procedures, undefined procedure scope, incorrect procedure name (length > 30 and format <> ^([A-Z][a-zA-Z]([a-ZA-Z0-9]){1,28})$)
' - Checking for procedures with more than 30 lines
' - Cutting lines upper than X characters
' -- Checking for non-used var
' --- Checking for code duplication
Public Function FormatModule(ByVal moduleName As String, Optional ByVal spaceTab As Integer = 4, Optional ByVal maxLen As Integer = 200) As String
    Dim i As Integer

    tabSpaces = ""
    For i = 1 To spaceTab
        tabSpaces = tabSpaces & " "
    Next i

    FormatModule = FormatModule & CheckModuleName(moduleName)
    Call WrapLines(moduleName, maxLen)
    Call SortModuleProc(moduleName)
    Call IndentCorrection(moduleName)
    Call CommentCorrection(moduleName)
    Call BlankLineCorrection(moduleName)
    Call OrganizeVar(moduleName)
    Call CorrectNext(moduleName)
    FormatModule = CheckProcedureNames(moduleName)
    FormatModule = FormatModule & CheckProcLength(moduleName)
End Function

' Function to format var declaration
Private Function FormatVar(ByVal codeLine As String) As String
    Dim i As Long
    Dim elem() As String, txt As String, chaine As String
    Dim arr As Variant, sortedVar As Variant, varDef As Variant, el As Variant

    If (tabSpaces = "") Then
        tabSpaces = "    "
    End If

    FormatVar = ""
    arr = Split(Mid(codeLine, 2), ",")
    sortedVar = Array()

    For Each varDef In arr
        If (InStr(varDef, " As ") = 0) Then
            varDef = CStr(varDef) & " As Variant"
        End If
        elem = Split(varDef, " As ")
        If (Not Contains(sortedVar, Trim(CStr(elem(1))))) Then
            Call AddArray(sortedVar, Trim(CStr(elem(1))))
        End If
    Next varDef
    Call QuickSort(sortedVar)

    chaine = ""
    For Each varDef In sortedVar
        txt = tabSpaces & "Dim "
        For i = 0 To UBound(arr)
            If (Trim(CStr(Split(arr(i), " As ")(1))) = Trim(CStr(varDef))) Then
                txt = txt & Trim(CStr(arr(i))) & ","
            End If
        Next i
        chaine = chaine & Mid(txt, 1, Len(txt) - 1) & Chr(13)
    Next varDef
    FormatVar = Mid(chaine, 1, Len(chaine) - 1)
End Function

' Function to get procedure name from code line
Private Function GetProcedureName(ByVal codeLine As String) As String
    Dim procName As String

    procName = Replace(codeLine, "Private ", "")
    procName = Replace(procName, "Public ", "")
    procName = Replace(procName, "Friend ", "")
    procName = Replace(procName, "Sub ", "")
    procName = Replace(procName, "Function ", "")
    GetProcedureName = Split(procName, "(")(0)
End Function

' Function to correct indentation of module
Public Sub IndentCorrection(ByVal moduleName As String)
    Dim codeMod As CodeModule
    Dim count As Integer, prevCount As Integer
    Dim i As Long
    Dim codeLine As String, curLine As String, prevLine As String, nextLine As String

    Set codeMod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
    
    With codeMod
        If (.CountOfLines = 0) Then
            Exit Sub
        End If
        
        For i = 1 To .CountOfDeclarationLines
            .ReplaceLine i, Trim(.Lines(i, 1))
        Next i

        .ReplaceLine i, CorrectLineIndent(curLine, 0)
        count = 1
        For i = .CountOfLines - 1 To .CountOfDeclarationLines + 1 Step -1
            prevCount = count
            curLine = Trim(.Lines(i, 1))
            prevLine = Trim(.Lines(i - 1, 1))
            nextLine = Trim(.Lines(i + 1, 1))
            If (EndWith(curLine, " _") And Not EndWith(nextLine, " _")) Then
                .ReplaceLine i + 1, CorrectLineIndent(nextLine, count + 1)
                count = prevCount
            End If
            If (Not EndWith(curLine, " _")) Then
                If (EndWith(prevLine, " _")) Then
                    count = count + 1
                End If
            ElseIf (EndWith(curLine, " _") And Not EndWith(prevLine, " _")) Then
                count = count - 2
            End If
            .ReplaceLine i, CorrectLineIndent(curLine, count)
        Next i
    End With
End Sub

' Function to check if line is var declaration (Dim and ReDim only)
Private Function IsDim(ByVal codeLine As String) As Boolean
    IsDim = StartWithList(codeLine, Array("Dim ", "ReDim "))
End Function

' Function to check if a string is a procedure declaration
Private Function IsProcedure(ByVal codeLine As String) As Boolean
    IsProcedure = (StartWithList(codeLine, Array("Public Function ", "Private Function ", "Function ", "Public Sub ", "Private Sub ", "Sub ", "Friend ", _
            "Private Declare ", "Public Declare ", "Declare ")))
End Function

' Function to check if line is a tab label
Private Function IsTabLabel(ByVal line As String) As Boolean
    IsTabLabel = (InStr(line, " ") = 0 And EndWith(line, ":"))
End Function

' Function to check if codeLine is For on one line
Private Function OneLineFor(ByVal codeLine As String) As Boolean
    OneLineFor = (StartWith(codeLine, "For ") And (InStr(codeLine, ":") <> 0))
End Function

' Function to check if codeLine is If Then on one line
Private Function OneLineIf(ByVal codeLine As String) As Boolean
    OneLineIf = (StartWith(codeLine, "If ") And Not EndWith(codeLine, "Then"))
End Function

' Function to reorganize var in a module
Public Sub OrganizeVar(ByVal moduleName As String)
    Dim codeMod As CodeModule
    Dim i As Long, j As Long, k As Long, moduleEnd As Long
    Dim codeLine As String

    Set codeMod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule

    With codeMod
        moduleEnd = .CountOfLines
        For i = 1 To moduleEnd
            codeLine = ""
            j = i
            k = i
            While (IsDim(Trim(.Lines(j, 1))) And Not StartWith(Trim(.Lines(j, 1)), "ReDim "))
                codeLine = codeLine & Replace(Trim(.Lines(j, 1)), "Dim ", ",")
                j = j + 1
                i = j + 1
            Wend
            If (k <> j) Then
                .DeleteLines k, j - k
                .InsertLines k, FormatVar(codeLine)
                moduleEnd = .CountOfLines
            End If
        Next i
    End With
End Sub

' Function to get position of first letter in string (returns 0 if there is no letter)
Private Function PosFirstLetter(ByVal str As String) As Long
    Dim chara As Integer

    PosFirstLetter = 1
    While (PosFirstLetter <= Len(str))
        chara = asc(Mid(str, PosFirstLetter, 1))
        If ((chara >= 65 And chara <= 90) Or (chara >= 97 And chara <= 122)) Then Exit Function
        PosFirstLetter = PosFirstLetter + 1
    Wend
    PosFirstLetter = 0
End Function

' Function to get all procedure names within a module as an Array
Public Function ProcNamesModule(ByVal moduleName As String) As Variant
    Dim codeMod As CodeModule
    Dim startLine As Long
    Dim procName As String
    Dim tmpArray As Variant

    Set codeMod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule

    tmpArray = Array()
    With codeMod
        startLine = .CountOfDeclarationLines + 1
        While startLine < .CountOfLines
            Call AddArray(tmpArray, .ProcOfLine(startLine, vbext_pk_Proc))
            startLine = startLine + .ProcCountLines(.ProcOfLine(startLine, vbext_pk_Proc), vbext_pk_Proc)
        Wend
    End With
    ProcNamesModule = tmpArray
End Function

' Quick sort
Private Sub QuickSort(ByRef arr As Variant)
    Call QuickSortRecursive(arr, 0, UBound(arr))
End Sub

' Recursive quick sort
Private Sub QuickSortRecursive(ByRef arr As Variant, ByVal leftIndex As Variant, ByVal rightIndex As Variant)
    Dim i As Variant, j As Variant, tmp As Variant, pivot As Variant

    i = leftIndex
    j = rightIndex

    pivot = arr((leftIndex + rightIndex) / 2)

    Do
        While pivot > arr(i)
            i = i + 1
        Wend
        While arr(j) > pivot
            j = j - 1
        Wend

        If j + 1 > i Then
            tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            j = j - 1: i = i + 1
        End If
    Loop Until i > j

    If leftIndex < j Then Call QuickSortRecursive(arr, leftIndex, j)
    If i < rightIndex Then Call QuickSortRecursive(arr, i, rightIndex)
End Sub

' Function to sort (delete and add) all procedures within a module (including comments) sorted by lexicographical order
Public Sub SortModuleProc(ByVal moduleName As String)
    Dim codeMod As CodeModule
    Dim procTxt As Variant, sortedProc As Variant

    Set codeMod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
    sortedProc = SortedModuleProc(moduleName)
    With codeMod
        .DeleteLines .CountOfDeclarationLines + 1, (.CountOfLines - .CountOfDeclarationLines)
        For Each procTxt In sortedProc
            .InsertLines .CountOfLines + 1, procTxt
        Next procTxt
    End With
End Sub

' Function to get all procedures within a module (including comments) sorted by lexicographical order
Public Function SortedModuleProc(ByVal moduleName As String) As Variant
    Dim i As Long
    Dim tmpArray As Variant

    tmpArray = ProcNamesModule(moduleName)
    Call QuickSort(tmpArray)
    For i = 0 To UBound(tmpArray)
        tmpArray(i) = CopyProcedure(tmpArray(i), moduleName)
    Next i
    SortedModuleProc = tmpArray
End Function

' Function to check if string starts with other string (str starts with start) (no trim)
Private Function StartWith(ByVal str As String, ByVal start As String, Optional ByVal withCase As Boolean = True) As Boolean
    StartWith = IIf(withCase, (Mid(str, 1, Len(start)) = start), (Mid(UCase(str), 1, Len(start)) = UCase(start)))
End Function

' Function to check if a string start with one of the elements in array
Private Function StartWithList(ByVal str As String, ByVal arr As Variant, Optional ByVal withCase As Boolean = True, Optional ByVal toTrim As Boolean = False) As Boolean
    Dim elem As Variant

    StartWithList = False
    
    If (toTrim) Then
        str = Trim(str)
    End If
    
    For Each elem In arr
        If (StartWith(str, CStr(elem), withCase)) Then
            StartWithList = True
            Exit Function
        End If
    Next elem
End Function

' Function to check if code line uses procedure or var
Private Function UsedCode(ByVal codeLine As String, ByVal codeName As String, Optional ByVal isProc As Boolean = True) As Boolean
    UsedCode = ((InStr(codeLine, codeName) > 0 And Not StartWith(codeLine, "'")) And ((isProc And (Not IsProcedure(codeLine))) Or (Not IsDim(codeLine))))
End Function

' Function to loop through module to check for procedure use
Private Function UsedInModule(ByVal projComp As VBComponent, ByVal procedureName As String, Optional ByVal isPrivate As Boolean = True, Optional ByVal moduleName As String = "") As Boolean
    Dim modulecode As CodeModule
    Dim j As Long
    Dim trimLine As String

    UsedInModule = False
    Set modulecode = projComp.CodeModule

    With modulecode
        If (isPrivate) Then
            For j = .CountOfDeclarationLines + 1 To .CountOfLines
                trimLine = Trim(.Lines(j, 1))
                If (UsedCode(trimLine, procedureName, True) And (procedureName <> .ProcOfLine(j, vbext_pk_Proc))) Then
                    UsedInModule = True
                    Exit Function
                End If
            Next j
        Else
            For j = .CountOfDeclarationLines + 1 To .CountOfLines
                trimLine = Trim(.Lines(j, 1))
                If (UsedCode(trimLine, procedureName, True) And ((procedureName <> modulecode.ProcOfLine(j, vbext_pk_Proc)) _
                                Or ((procedureName <> modulecode.ProcOfLine(j, vbext_pk_Proc)) And (projComp.Name <> moduleName)))) Then
                    UsedInModule = True
                    Exit Function
                End If
            Next j
        End If
    End With
End Function

' Function to wrap a line at last space before maxLen character
Private Function WrapLine(ByVal codeLine As String, ByVal maxLen As Long) As String
    Dim countString As Integer
    Dim i As Long

    If Len(codeLine) > maxLen Then
        i = 0
Top:
        If i = maxLen Then
            WrapLine = codeLine
            Exit Function
        End If
    
        If (Mid(codeLine, maxLen - i, 1) = " ") Then
            WrapLine = IIf(CountOccurences(Mid(codeLine, 1, maxLen - i), Chr(34)) Mod 2 = 1, Mid(codeLine, 1, maxLen - i) & Chr(34) & " & _ " & Chr(13) & Chr(34) & Mid(codeLine, maxLen - i + 1), _
                    Mid(codeLine, 1, maxLen - i) & "_ " & Chr(13) & Mid(codeLine, maxLen - i + 1))
        Else
            i = i + 1
            GoTo Top
        End If
    Else
        WrapLine = codeLine
    End If
End Function

' Function to wrap lines upper than x characters in module
Public Sub WrapLines(ByVal moduleName As String, Optional ByVal maxLen As Integer = 200)
    Dim codeMod As CodeModule
    Dim j As Integer
    Dim i As Long
    Dim codeLine As String, wrappedLines() As String, splitLine() As String

    Set codeMod = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule

    With codeMod
        For i = .CountOfLines To 1 Step -1
            codeLine = .Lines(i, 1)
            If (Len(codeLine) > maxLen) Then
                ReDim wrappedLines(Int(Len(codeLine) / maxLen) + 1)
                For j = 0 To UBound(wrappedLines)
                    splitLine = Split(WrapLine(codeLine, maxLen), Chr(13))
                    wrappedLines(j) = splitLine(0)
                    If (UBound(splitLine) > 0) Then
                        codeLine = AddTabs(splitLine(1), CountTabulations(wrappedLines(0)) + 2)
                    Else
                        Exit For
                    End If
                Next j
                .ReplaceLine i, Mid(Join(wrappedLines, Chr(13)), 1, Len(Join(wrappedLines, Chr(13))) - 1)
            End If
        Next i
    End With
End Sub