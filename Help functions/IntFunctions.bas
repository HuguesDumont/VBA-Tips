Attribute VB_Name = "intFunctions"
Attribute VB_Description = " Pre-made functions and subs using Integer type\n"
Option Explicit

'Function to generate an integer (int) between 2 values (min, max)
Public Function intRndBetween(min As Integer, max As Integer) As Integer
Attribute intRndBetween.VB_Description = "Function to generate an integer (int) between 2 values (min, max)"
    Randomize
    intRndBetween = Int((max - min + 1) * Rnd + min)
End Function

'Sub to sort an array of integer integer ascending
'Needs the "intQuickSortAsc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub intSortAsc(ByRef arr() As Integer)
Attribute intSortAsc.VB_Description = "Sub to sort an array of integer integer ascending\r\nNeeds the ""intQuickSortAsc"" sub to work\r\nDefault ubound for array is 0 and max index is ubound-1"
    Call intQuickSortAsc(arr, 0, UBound(arr))
End Sub

'Sub to sort an array of integer integer descending
'Needs the "intQuickSortDesc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub intSortDesc(ByRef arr() As Integer)
Attribute intSortDesc.VB_Description = "Sub to sort an array of integer integer descending\r\nNeeds the ""intQuickSortDesc"" sub to work\r\nDefault ubound for array is 0 and max index is ubound-1"
    Call intQuickSortDesc(arr, 0, UBound(arr))
End Sub

'Recursive quicksort for ascending sort of integer int array
Public Sub intQuickSortAsc(ByRef arr() As Integer, ByVal leftIndex As Integer, ByVal rightIndex As Integer)
Attribute intQuickSortAsc.VB_Description = "Recursive quicksort for ascending sort of integer int array"
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
        
    If leftIndex < j Then Call intQuickSortAsc(arr(), leftIndex, j)
    If i < rightIndex Then Call intQuickSortAsc(arr(), i, rightIndex)
End Sub

'Recursive quicksort for descending sort of integer int array
Public Sub intQuickSortDesc(arr() As Integer, leftIndex As Integer, rightIndex As Integer)
Attribute intQuickSortDesc.VB_Description = "Recursive quicksort for descending sort of integer int array"
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

    If (leftIndex < j) Then intQuickSortDesc arr, leftIndex, j
    If (i < rightIndex) Then intQuickSortDesc arr, i, rightIndex
End Sub

'Function to sum all values in integer array (!CARERUL! returns a long in case of overflowing capacity)
Public Function sumIntArray(arr() As Integer) As Long
Attribute sumIntArray.VB_Description = "Function to sum all values in integer array (!CARERUL! returns a long in case of overflowing capacity)"
    Dim i As Long
    
    sumIntArray = 0
    For i = 0 To UBound(arr)
        sumIntArray = sumIntArray + arr(i)
    Next i
End Function

'Function to calculate average value of integer array (!CARERUL! returns result as double)
'Works with the "sumIntArray" from IntFunctions
Public Function averageInt(arr() As Integer) As Double
Attribute averageInt.VB_Description = "Function to calculate average value of integer array (!CARERUL! returns result as double)\r\nWorks with the ""sumIntArray"" from IntFunctions"
    averageInt = CDbl(sumIntArray(arr)) / CDbl(UBound(arr))
End Function

'Function to return the max of Integer array
Public Function maxInteger(arr() As Integer) As Integer
Attribute maxInteger.VB_Description = "Function to return the max of Integer array"
    Dim i As Long
    
    maxInteger = arr(0)
    For i = 1 To UBound(arr)
        If maxInteger < arr(i) Then maxInteger = arr(i)
    Next i
End Function

'Function to return the min of Integer array
Public Function minInteger(arr() As Integer) As Integer
Attribute minInteger.VB_Description = "Function to return the min of Integer array"
    Dim i As Long
    
    minInteger = arr(0)
    For i = 1 To UBound(arr)
        If minInteger > arr(i) Then minInteger = arr(i)
    Next i
End Function

'Function to find an int value in int array (returns a Long : -1 if not found, position in array if found)
Public Function findInt(arr() As Integer, val As Integer) As Long
Attribute findInt.VB_Description = "Function to find an int value in int array (returns a Long : -1 if not found, position in array if found)"
    Dim i As Long
    
    findInt = -1
    For i = 0 To UBound(arr)
        If arr(i) = val Then
            findInt = i
            Exit Function
        End If
    Next i
End Function
