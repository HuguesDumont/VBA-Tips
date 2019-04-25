Attribute VB_Name = "ContextualMenu"
'Need the "Microsoft Forms 2.0 Object Library" reference activated to work
Option Explicit

#If Mac Then
    ' do nothing
#Else
    #If VBA7 Then
        Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
        Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
        Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
        Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
        Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As LongPtr
        Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long
        Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
        Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    #Else
        Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
        Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
        Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
        Declare Function CloseClipboard Lib "User32" () As Long
        Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
        Declare Function EmptyClipboard Lib "User32" () As Long
        Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
        Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    #End If
#End If

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Public param As UserForm

'Function to set clipboard data
Sub ClipBoardSetData(MyString As String)
    Dim x As Long
    
    #If Mac Then
        With New MSForms.DataObject
            .SetText MyString
            .PutInClipboard
        End With
    #Else
        #If VBA7 Then
            Dim hGlobalMemory As LongPtr
            Dim hClipMemory As LongPtr
            Dim lpGlobalMemory As LongPtr
        #Else
            Dim hGlobalMemory As Long
            Dim hClipMemory As Long
            Dim lpGlobalMemory As Long
        #End If

        ' Allocate moveable global memory.
        '-------------------------------------------
        hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

        ' Lock the block to get a far pointer to this memory.
        lpGlobalMemory = GlobalLock(hGlobalMemory)

        ' Copy the string to this global memory.
        lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

        ' Unlock the memory.
        If GlobalUnlock(hGlobalMemory) <> 0 Then
            MsgBox "Error while trying to copy to clipboard."
            Exit Sub
        End If

        ' Open the Clipboard to copy data to.
        If OpenClipboard(0&) = 0 Then
            MsgBox "Error while trying to copy to clipboard."
            Exit Sub
        End If

        ' Clear the Clipboard.
        x = EmptyClipboard()

        ' Copy the data to the Clipboard.
        hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)
        
        'Do not forget to close clipboard.
        Call CloseClipboard
    #End If
End Sub

'Creates the contextual menu
Sub CreateContextual()
    Dim Barre As CommandBar
    Dim Controle As CommandBarControl
    
    On Error Resume Next
    CommandBars.Item("Contextual").Delete
    Set Barre = CommandBars.Add(name:="Contextual", Position:=msoBarPopup, temporary:=True)
    
    Set Controle = Barre.Controls.Add
    With Controle
        .Caption = "Cut"
        .OnAction = "mnCut"
    End With
    
    Set Controle = Barre.Controls.Add
    With Controle
        .Caption = "Copy"
        .OnAction = "mnCopy"
    End With
    
    Set Controle = Barre.Controls.Add
    With Controle
        .Caption = "Paste"
        .OnAction = "mnPaste"
    End With
    
    Set Controle = Barre.Controls.Add
    With Controle
        .Caption = "To lower"
        .OnAction = "mnLower"
    End With
    
    Set Controle = Barre.Controls.Add
    With Controle
        .Caption = "To upper"
        .OnAction = "mnUpper"
    End With
 
    Set Controle = Barre.Controls.Add
    With Controle
        .Caption = "Clear selection"
        .OnAction = "mnSelectClear"
    End With
    
    Set Controle = Barre.Controls.Add
    With Controle
        .Caption = "Clear content"
        .OnAction = "mnClearContent"
    End With
    
    Set Controle = Barre.Controls.Add
    With Controle
        .Caption = "Select all"
        .OnAction = "mnSelectAll"
    End With
    
    Set Controle = Nothing
    Set Barre = Nothing
End Sub

'Cut menu
Sub mnCut()
    Dim data As New DataObject
    
    On Error Resume Next
    
    With param.ActiveControl
        data.SetText .SelText
        .SelText = ""
        data.PutInClipboard
    End With
End Sub

'Copy menu
Sub mnCopy()
    Dim data As New DataObject
    
    On Error Resume Next
    
    ClipBoardSetData param.ActiveControl.SelText
    Set data = Nothing
End Sub

'Paste menu
Sub mnPaste()
    Dim data As New DataObject

    On Error Resume Next
    
    data.GetFromClipboard
    param.ActiveControl.SelText = data.GetText
End Sub

'Lower text menu
Sub mnLower()
    On Error Resume Next
    param.ActiveControl.SelText = LCase(param.ActiveControl.SelText)
End Sub

'Upper text menu
Sub mnUpper()
    On Error Resume Next
    param.ActiveControl.SelText = UCase(param.ActiveControl.SelText)
End Sub

'Clear selection menu
Sub mnSelectClear()
    On Error Resume Next
    param.ActiveControl.SelText = ""
End Sub

'Clear all content menu
Sub mnClearContent()
    On Error Resume Next
    param.ActiveControl.Value = ""
End Sub

'Select all menu
Sub mnSelectAll()
    On Error Resume Next
    param.ActiveControl.SelStart = 0
    param.ActiveControl.SelLength = Len(param.ActiveControl.Value)
End Sub


