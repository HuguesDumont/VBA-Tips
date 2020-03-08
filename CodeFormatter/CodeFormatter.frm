VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CodeFormatter 
   Caption         =   "Code Formatter by Hugues DUMONT"
   ClientHeight    =   10470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14550
   OleObjectBlob   =   "CodeFormatter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CodeFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Made by Hugues DUMONT
' This project is still in development, therefore, there is only a basis of what the future project will be like

Private Sub ChangeLength(ByRef ctrl As Control, ByVal up As Boolean, Optional ByVal limit As Long = 0)
    If (ctrl.value <> "") Then
        ctrl.value = IIf(up, IIf(CLng(ctrl.value) < limit, ctrl.value + 1, limit), IIf(CLng(ctrl.value) > 0, ctrl.value - 1, 0))
    End If
End Sub

Private Function CheckLimitValue(ByVal strVal As String, ByVal limit As Long) As Boolean
    Dim reg As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^(([0])|(([1-9])([\d]){0,4}))$"
    
    If (reg.test(strVal)) Then
        CheckLimitValue = (CLng(strVal) <= limit)
    Else
        CheckLimitValue = False
    End If
    
    Set reg = Nothing
End Function

Private Sub FormatCodeButton_Click()
    
End Sub

Private Sub LineLength_Change()
    Me.LineLength.value = IIf(CheckLimitValue(Me.LineLength.value, 250), Me.LineLength.value, 200)
End Sub

Private Sub LineSpin_SpinDown()
    Call ChangeLength(Me.LineLength, False)
End Sub

Private Sub LineSpin_SpinUp()
    Call ChangeLength(Me.LineLength, True, 250)
End Sub

Private Sub ListElements_Change()
    Dim procArray As Variant
    Dim i As Integer
    
    procArray = Array()
    
    For i = 0 To Me.ListElements.ListCount - 1
        If (Me.ListElements.Selected(i)) Then
            Call AddAllArray(GetSortedProcedures(Me.ListElements.List(i)), procArray)
        End If
    Next i
    
    Me.ListProcedures.Enabled = True
    Me.ListProcedures.List = procArray
End Sub

Private Sub ModuleLength_Change()
    Me.ModuleLength.value = IIf(CheckLimitValue(Me.ModuleLength.value, 10000), Me.ModuleLength.value, 1000)
End Sub

Private Sub ModuleSpin_SpinDown()
    Call ChangeLength(Me.ModuleLength, False)
End Sub

Private Sub ModuleSpin_SpinUp()
    Call ChangeLength(Me.ModuleLength, True, 10000)
End Sub

Private Sub OptionClass_Click()
    Me.ListElements.List = GetSortedComponents(3)
    Me.ListElements.Enabled = False
    Me.ListProcedures.Enabled = False
    Me.OptionProcedures.value = False
    Me.OptionProcedures.Enabled = False
End Sub

Private Sub OptionDocuments_Click()
    Me.ListElements.List = GetSortedComponents(100)
    Me.ListElements.Enabled = False
    Me.ListProcedures.Enabled = False
    Me.OptionProcedures.value = False
    Me.OptionProcedures.Enabled = False
End Sub

Private Sub OptionModule_Click()
    Me.ListElements.List = GetSortedComponents(1)
    Me.ListElements.Enabled = False
    Me.ListProcedures.Enabled = False
    Me.OptionProcedures.value = False
    Me.OptionProcedures.Enabled = False
End Sub

Private Sub OptionProcedures_Click()
    Dim procArray As Variant
    Dim i As Integer
    
    procArray = Array()
    
    For i = 0 To Me.ListElements.ListCount - 1
        If (Me.ListElements.Selected(i)) Then
            Call AddAllArray(GetSortedProcedures(Me.ListElements.List(i)), procArray)
        End If
    Next i
    
    Me.ListProcedures.Enabled = True
    Me.ListProcedures.List = procArray
End Sub

Private Sub OptionProject_Click()
    Me.ListElements.List = GetSortedComponents()
    Me.ListElements.Enabled = False
    Me.ListProcedures.Enabled = False
    Me.OptionProcedures.value = False
    Me.OptionProcedures.Enabled = False
End Sub

Private Sub OptionSpecific_Click()
    Me.ListElements.List = GetSortedComponents()
    Me.ListElements.Enabled = True
    Me.ListProcedures.Enabled = False
    Me.OptionProcedures.Enabled = True
End Sub

Private Sub ProcedureLength_Change()
    Me.ProcedureLength.value = IIf(CheckLimitValue(Me.ProcedureLength.value, 100), Me.ProcedureLength.value, 30)
End Sub

Private Sub ProcedureSpin_SpinDown()
    Call ChangeLength(Me.ProcedureLength, False)
End Sub

Private Sub ProcedureSpin_SpinUp()
    Call ChangeLength(Me.ProcedureLength, True, 100)
End Sub

Private Sub SelectAll_Click()
    Dim ctrl As Control
    
    For Each ctrl In Me.FormatOptions.Controls
        If TypeName(ctrl) = "CheckBox" Then
            ctrl.value = True
        End If
    Next ctrl
End Sub

Private Sub TabLength_Change()
    Me.TabLength.value = IIf(CheckLimitValue(Me.TabLength.value, 10), Me.TabLength.value, "")
End Sub

Private Sub TabSpin_SpinDown()
    Call ChangeLength(Me.TabLength, False)
End Sub

Private Sub TabSpin_SpinUp()
    Call ChangeLength(Me.TabLength, True, 10)
End Sub

Private Sub UnselectAll_Click()
    Dim ctrl As Control
    
    For Each ctrl In Me.FormatOptions.Controls
        If TypeName(ctrl) = "CheckBox" Then
            ctrl.value = False
        End If
    Next ctrl
End Sub

Private Sub UserForm_Initialize()
    Me.ListElements.List = GetSortedComponents()
End Sub

Private Function GetSortedComponents(Optional ByVal compType As Integer = -1) As Variant
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim elemArr As Variant
    
    Set vbProj = ThisWorkbook.VBProject
    
    elemArr = Array()
    
    For Each vbComp In vbProj.VBComponents
        If (compType = -1 Or vbComp.Type = compType) Then
            Call AddArray(vbComp.Name, elemArr)
        End If
    Next vbComp
    
    Call QuickSort(elemArr)
    GetSortedComponents = elemArr
End Function

Private Function GetSortedProcedures(ByVal moduleName As String) As Variant
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim codeMod As VBIDE.CodeModule
    Dim procKind As VBIDE.vbext_ProcKind
    Dim procArray As Variant
    Dim line As Integer
    Dim procName As String
    
    Set vbProj = ThisWorkbook.VBProject
    Set vbComp = vbProj.VBComponents(moduleName)
    Set codeMod = vbComp.CodeModule
    
    procArray = Array()
    
    With codeMod
        For line = .CountOfDeclarationLines + 1 To .CountOfLines - 1
            procName = .ProcOfLine(line, procKind)
            Call AddArray(procName, procArray)
            line = .ProcStartLine(procName, procKind) + .ProcCountLines(procName, procKind)
        Next line
    End With
    GetSortedProcedures = procArray
End Function
