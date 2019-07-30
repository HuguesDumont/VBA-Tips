Attribute VB_Name = "intFunctions"
Attribute VB_Description = " Pre-made functions and subs using Integer type\n"
Option Explicit

'Function to generate an integer (int) between 2 values (min, max)
Public Function IntRndBetween(ByVal min As Integer, ByVal max As Integer) As Integer
Attribute IntRndBetween.VB_Description = "Function to generate an integer (int) between 2 values (min, max)"
    Randomize
    IntRndBetween = Int((max - min + 1) * Rnd + min)
End Function

'Sub to sort an array of integer integer ascending
'Needs the "intQuickSortAsc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub IntSortAsc(ByRef arr() As Integer)
Attribute IntSortAsc.VB_Description = "Sub to sort an array of integer integer ascending\r\nNeeds the ""intQuickSortAsc"" sub to work\r\nDefault ubound for array is 0 and max index is ubound-1"
    Call IntQuickSortAsc(arr, 0, UBound(arr))
End Sub

'Sub to sort an array of integer integer descending
'Needs the "intQuickSortDesc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub IntSortDesc(ByRef arr() As Integer)
Attribute IntSortDesc.VB_Description = "Sub to sort an array of integer integer descending\r\nNeeds the ""intQuickSortDesc"" sub to work\r\nDefault ubound for array is 0 and max index is ubound-1"
    Call IntQuickSortDesc(arr, 0, UBound(arr))
End Sub

'Recursive quicksort for ascending sort of integer int array
Public Sub IntQuickSortAsc(ByRef arr() As Integer, ByVal leftIndex As Integer, ByVal rightIndex As Integer)
Attribute IntQuickSortAsc.VB_Description = "Recursive quicksort for ascending sort of integer int array"
    Dim i As Integer, j As Integer, tmp As Integer, pivot As Integer
    
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
        
    If leftIndex < j Then Call IntQuickSortAsc(arr(), leftIndex, j)
    If i < rightIndex Then Call IntQuickSortAsc(arr(), i, rightIndex)
End Sub

'Recursive quicksort for descending sort of integer int array
Public Sub IntQuickSortDesc(ByRef arr() As Integer, ByVal leftIndex As Integer, ByVal rightIndex As Integer)
Attribute IntQuickSortDesc.VB_Description = "Recursive quicksort for descending sort of integer int array"
    Dim pivot As Integer, tmp As Integer, i As Integer, j As Integer

    i = leftIndex
    j = rightIndex

    pivot = arr((leftIndex + rightIndex) \ 2)

    While (i <= j)
        While (arr(i) > pivot And i < rightIndex) 'converted sign
            i = i + 1
        Wend

        While (pivot > arr(j) And j > leftIndex) 'converted sign
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

'Function to sum all values in integer array (!CARERUL! returns a long in case of overflowing capacity)
Public Function SumIntArray(ByRef arr() As Integer) As Long
Attribute SumIntArray.VB_Description = "Function to sum all values in integer array (!CARERUL! returns a long in case of overflowing capacity)"
    Dim i As Long
    
    SumIntArray = 0
    For i = 0 To UBound(arr)
        SumIntArray = SumIntArray + arr(i)
    Next i
End Function

'Function to calculate average value of integer array (!CARERUL! returns result as double)
'Works with the "sumIntArray" from IntFunctions
Public Function AverageInt(ByRef arr() As Integer) As Double
Attribute AverageInt.VB_Description = "Function to calculate average value of integer array (!CARERUL! returns result as double)\r\nWorks with the ""sumIntArray"" from IntFunctions"
    AverageInt = CDbl(SumIntArray(arr)) / CDbl(UBound(arr))
End Function

'Function to return the max of Integer array
Public Function MaxInteger(ByRef arr() As Integer) As Integer
Attribute MaxInteger.VB_Description = "Function to return the max of Integer array"
    Dim i As Long
    
    MaxInteger = arr(0)
    For i = 1 To UBound(arr)
        If MaxInteger < arr(i) Then MaxInteger = arr(i)
    Next i
End Function

'Function to return the min of Integer array
Public Function MinInteger(ByRef arr() As Integer) As Integer
Attribute MinInteger.VB_Description = "Function to return the min of Integer array"
    Dim i As Long
    
    MinInteger = arr(0)
    For i = 1 To UBound(arr)
        If MinInteger > arr(i) Then MinInteger = arr(i)
    Next i
End Function

'Function to find an int value in int array (returns a Long : -1 if not found, position in array if found)
Public Function FindInt(ByRef arr() As Integer, ByVal val As Integer) As Long
Attribute FindInt.VB_Description = "Function to find an int value in int array (returns a Long : -1 if not found, position in array if found)"
    Dim i As Long
    
    FindInt = -1
    For i = 0 To UBound(arr)
        If arr(i) = val Then
            FindInt = i
            Exit Function
        End If
    Next i
End Function
