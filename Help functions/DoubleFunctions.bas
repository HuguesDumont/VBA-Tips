Attribute VB_Name = "DoubleFunctions"
Attribute VB_Description = "Pre-made functions and subs using double"
Option Explicit

'Function to generate a double between 2 values (min, max)
Public Function doubleRndBetween(min As Double, max As Double) As Double
Attribute doubleRndBetween.VB_Description = "Function to generate a double between 2 values (min, max)"
    Dim rndVariable As Double
    
    Randomize
    rndVariable = Rnd
    
    If (max - min + 1) * rndVariable + min <= max Then
        doubleRndBetween = (max - min + 1) * rndVariable + min
    Else
        Do While (max - min + 1) * rndVariable + min > max
            rndVariable = Rnd
            If (max - min + 1) * rndVariable + min <= max Then
                doubleRndBetween = (max - min + 1) * rndVariable + min
            End If
        Loop
    End If
End Function

'Sub to sort an array of double ascending
'Needs the "doubleQuickSortAsc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub doubleSortAsc(ByRef arr() As Double)
Attribute doubleSortAsc.VB_Description = "test"
    Call doubleQuickSortAsc(arr, 0, UBound(arr) - 1)
End Sub

'sub to sort an array of double descending
'Needs the "doubleQuickSortDesc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub doubleSortDesc(ByRef arr() As Double)
    Call doubleQuickSortDesc(arr, 0, UBound(arr) - 1)
End Sub

'recursive quick sort for ascending sort of double array
Public Sub doubleQuickSortAsc(ByRef arr() As Double, ByVal leftIndex As Double, ByVal rightIndex As Double)
    Dim i As Double, j As Double, tmp As Double, pivot As Double
    
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
        
    If leftIndex < j Then Call doubleQuickSortAsc(arr(), leftIndex, j)
    If i < rightIndex Then Call doubleQuickSortAsc(arr(), i, rightIndex)
End Sub

'recursive quick sort for descending sort of double array
Public Sub doubleQuickSortDesc(arr() As Double, leftIndex As Double, rightIndex As Double)
    Dim pivot As Double, tmp As Double, i As Double, j As Double

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

    If (leftIndex < j) Then doubleQuickSortDesc arr, leftIndex, j
    If (i < rightIndex) Then doubleQuickSortDesc arr, i, rightIndex
End Sub

'Function to sum all values in double array
Public Function sumDoubleArray(arr() As Double) As Double
    Dim i As Long
    
    sumDoubleArray = 0
    For i = 0 To UBound(arr) - 1
        sumDoubleArray = sumDoubleArray + arr(i)
    Next i
End Function

'Function to calculate average value of double array
'Works with the "sumDoubleArray" from DoubleFunctions
Public Function averageDouble(arr() As Double) As Double
Attribute averageDouble.VB_Description = "test"
    averageDouble = sumDoubleArray(arr) / CDbl(UBound(arr))
End Function

'Function to return the max of Double array
Public Function maxDouble(arr() As Double) As Double
    Dim i As Long
    
    maxDouble = arr(0)
    For i = 1 To UBound(arr) - 1
        If maxDouble < arr(i) Then maxDouble = arr(i)
    Next i
End Function

'Function to return the min of Double array
Public Function minDouble(arr() As Double) As Double
    Dim i As Long
    
    minDouble = arr(0)
    For i = 1 To UBound(arr) - 1
        If minDouble > arr(i) Then minDouble = arr(i)
    Next i
End Function

'Function to find a double value in double array (returns a Long : -1 if not found, position in array if found)
Public Function findDouble(arr() As Double, val As Double) As Long
    Dim i As Long
    
    findDouble = -1
    For i = 0 To UBound(arr) - 1
        If arr(i) = val Then
            findDouble = i
            Exit Function
        End If
    Next i
End Function
