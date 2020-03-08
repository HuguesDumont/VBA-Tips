Attribute VB_Name = "IntFunctions"
Attribute VB_Description = " Pre-made functions and subs using Integer type"
Option Explicit

'Function to calculate average value of integer array (!CARERUL! returns result as double)
'Works with the "sumIntArray" from IntFunctions
Public Function AverageInt(ByRef arr() As Integer) As Double
    AverageInt = CDbl(SumIntArray(arr)) / CDbl(UBound(arr))
End Function

'Function to find an int value in int array (returns a Long : -1 if not found, position in array if found)
Public Function FindInt(ByRef arr() As Integer, ByVal val As Integer) As Long
    Dim i                               As Long
    
    FindInt = -1
    
    For i = 0 To UBound(arr)
        If (arr(i) = val) Then
            FindInt = i
            Exit Function
        End If
    Next i
End Function

'Recursive quicksort for ascending sort of integer int array
Public Sub IntQuickSortAsc(ByRef arr() As Integer, ByVal leftIndex As Integer, ByVal rightIndex As Integer)
    Dim i                               As Integer
    Dim j                               As Integer
    Dim tmp                             As Integer
    Dim pivot                           As Integer
    
    i = leftIndex
    j = rightIndex
    pivot = arr((leftIndex + rightIndex) / 2)
    
    Do
        While (pivot > arr(i))
            i = i + 1
        Wend
        
        While (arr(j) > pivot)
            j = j - 1
        Wend
        
        If ((j + 1) > i) Then
            tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            j = j - 1
            i = i + 1
        End If
    Loop Until (i > j)
    
    If (leftIndex < j) Then Call IntQuickSortAsc(arr(), leftIndex, j)
    If (i < rightIndex) Then Call IntQuickSortAsc(arr(), i, rightIndex)
End Sub

'Recursive quicksort for descending sort of integer int array
Public Sub IntQuickSortDesc(ByRef arr() As Integer, ByVal leftIndex As Integer, ByVal rightIndex As Integer)
    Dim pivot                           As Integer
    Dim tmp                             As Integer
    Dim i                               As Integer
    Dim j                               As Integer
    
    i = leftIndex
    j = rightIndex
    
    pivot = arr((leftIndex + rightIndex) \ 2)
    
    While (i <= j)
        While ((arr(i) > pivot) And (i < rightIndex)) 'converted sign
            i = i + 1
        Wend
        
        While ((pivot > arr(j)) And (j > leftIndex)) 'converted sign
            j = j - 1
        Wend
        
        If (i <= j) Then
            tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Wend
    
    If (leftIndex < j) Then IntQuickSortDesc arr, leftIndex, j
    If (i < rightIndex) Then IntQuickSortDesc arr, i, rightIndex
End Sub

'Function to generate an integer (int) between 2 values (min, max)
Public Function IntRndBetween(ByVal min As Integer, ByVal max As Integer) As Integer
    Randomize
    IntRndBetween = Int((max - min + 1) * Rnd + min)
End Function

'Sub to sort an array of integer integer ascending
'Needs the "intQuickSortAsc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub IntSortAsc(ByRef arr() As Integer)
    Call IntQuickSortAsc(arr, 0, UBound(arr))
End Sub

'Sub to sort an array of integer integer descending
'Needs the "intQuickSortDesc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub IntSortDesc(ByRef arr() As Integer)
    Call IntQuickSortDesc(arr, 0, UBound(arr))
End Sub

'Function to return the max of Integer array
Public Function MaxInteger(ByRef arr() As Integer) As Integer
    Dim i                               As Long
    
    MaxInteger = arr(0)
    
    For i = 1 To UBound(arr)
        If (MaxInteger < arr(i)) Then MaxInteger = arr(i)
    Next i
End Function

'Function to return the min of Integer array
Public Function MinInteger(ByRef arr() As Integer) As Integer
    Dim i                               As Long
    
    MinInteger = arr(0)
    
    For i = 1 To UBound(arr)
        If (MinInteger > arr(i)) Then MinInteger = arr(i)
    Next i
End Function

'Function to sum all values in integer array (!CARERUL! returns a long in case of overflowing capacity)
Public Function SumIntArray(ByRef arr() As Integer) As Long
    Dim i                               As Long
    
    SumIntArray = 0
    
    For i = 0 To UBound(arr)
        SumIntArray = SumIntArray + arr(i)
    Next i
End Function
