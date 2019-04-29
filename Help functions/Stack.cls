VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Implementation of simple stack"
Option Explicit

Private values As New Collection

'Get the number of items in Stack
Public Property Get Size() As Long
Attribute Size.VB_Description = "Get the number of items in Stack"
    Size = values.Count
End Property

'Get first item of the stack
Public Function Pop() As Variant
Attribute Pop.VB_Description = "Get first item of the stack"
    With values
        If .Count > 0 Then
            Pop = .item(.Count)
            .remove .Count
        End If
    End With
End Function

'Add an item to stack
Public Function Push(value As Variant)
Attribute Push.VB_Description = "Add an item to stack"
    values.Add value
End Function

'Check if stack is empty
Public Function IsEmpty() As Boolean
Attribute IsEmpty.VB_Description = "Check if stack is empty"
    IsEmpty = (values.Count = 0)
End Function
