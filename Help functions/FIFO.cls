VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FIFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Implementation of a FIFO (First In, First Out)"
Option Explicit

Private values As New Collection

'Get number of elements in FIFO
Public Property Get Size() As Long
Attribute Size.VB_Description = "Number of elements in FIFO"
    Size = values.count
End Property

'Get the first item of FIFO
Public Function Pop() As Variant
Attribute Pop.VB_Description = "Get the first item of FIFO"
    With values
        If .count > 0 Then
            Pop = .item(1)
            .remove 1
        End If
    End With
End Function

'Add an item to FIFO
Public Function Push(ByVal value As Variant)
Attribute Push.VB_Description = "Add an item to FIFO"
    values.Add value
End Function

'Check if FIFO is empty
Public Function IsEmpty() As Boolean
Attribute IsEmpty.VB_Description = "Function to check if FIFO is empty"
    IsEmpty = (values.count = 0)
End Function

