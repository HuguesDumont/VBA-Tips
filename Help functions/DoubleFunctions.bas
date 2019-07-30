Attribute VB_Name = "DoubleFunctions"
Attribute VB_Description = "Pre-made functions and subs using double type"
Option Explicit

'Function to generate a double between 2 values (min, max)
Public Function DoubleRndBetween(ByVal min As Double, ByVal max As Double) As Double
Attribute DoubleRndBetween.VB_Description = "Function to generate a double between 2 values (min, max)"
    Dim rndVariable As Double
    
    Randomize
    rndVariable = Rnd
    
    If (max - min + 1) * rndVariable + min <= max Then
        DoubleRndBetween = (max - min + 1) * rndVariable + min
    Else
        Do While (max - min + 1) * rndVariable + min > max
            rndVariable = Rnd
            If (max - min + 1) * rndVariable + min <= max Then
                DoubleRndBetween = (max - min + 1) * rndVariable + min
            End If
        Loop
    End If
End Function

'Sub to sort an array of double ascending
'Needs the "doubleQuickSortAsc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub DoubleSortAsc(ByRef arr() As Double)
Attribute DoubleSortAsc.VB_Description = "Sub to sort an array of double ascending\r\nNeeds the ""doubleQuickSortAsc"" sub to work\r\nDefault ubound for array is 0 and max index is ubound-1"
    Call DoubleQuickSortAsc(arr, 0, UBound(arr))
End Sub

'Sub to sort an array of double descending
'Needs the "doubleQuickSortDesc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub DoubleSortDesc(ByRef arr() As Double)
Attribute DoubleSortDesc.VB_Description = "Sub to sort an array of double descending\r\nNeeds the ""doubleQuickSortDesc"" sub to work\r\nDefault ubound for array is 0 and max index is ubound-1"
    Call DoubleQuickSortDesc(arr, 0, UBound(arr))
End Sub

'Recursive quick sort for ascending sort of double array
Public Sub DoubleQuickSortAsc(ByRef arr() As Double, ByVal leftIndex As Double, ByVal rightIndex As Double)
Attribute DoubleQuickSortAsc.VB_Description = "Recursive quick sort for ascending sort of double array"
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
        
    If leftIndex < j Then Call DoubleQuickSortAsc(arr(), leftIndex, j)
    If i < rightIndex Then Call DoubleQuickSortAsc(arr(), i, rightIndex)
End Sub

'Recursive quick sort for descending sort of double array
Public Sub DoubleQuickSortDesc(ByRef arr() As Double, ByVal leftIndex As Double, ByVal rightIndex As Double)
Attribute DoubleQuickSortDesc.VB_Description = "Recursive quick sort for descending sort of double array"
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

    If (leftIndex < j) Then DoubleQuickSortDesc arr, leftIndex, j
    If (i < rightIndex) Then DoubleQuickSortDesc arr, i, rightIndex
End Sub

'Function to sum all values in double array
Public Function SumDoubleArray(ByRef arr() As Double) As Double
Attribute SumDoubleArray.VB_Description = "Function to sum all values in double array"
    Dim i As Long
    
    SumDoubleArray = 0
    For i = 0 To UBound(arr)
        SumDoubleArray = SumDoubleArray + arr(i)
    Next i
End Function

'Function to calculate average value of double array
'Works with the "sumDoubleArray" from DoubleFunctions
Public Function AverageDouble(ByRef arr() As Double) As Double
Attribute AverageDouble.VB_Description = "Function to calculate average value of double array\r\nWorks with the ""sumDoubleArray"" from DoubleFunctions"
    AverageDouble = SumDoubleArray(arr) / CDbl(UBound(arr))
End Function

'Function to return the max of Double array
Public Function MaxDouble(ByRef arr() As Double) As Double
Attribute MaxDouble.VB_Description = "Function to return the max of Double array"
    Dim i As Long
    
    MaxDouble = arr(0)
    For i = 1 To UBound(arr)
        If MaxDouble < arr(i) Then MaxDouble = arr(i)
    Next i
End Function

'Function to return the min of Double array
Public Function MinDouble(ByRef arr() As Double) As Double
Attribute MinDouble.VB_Description = "Function to return the min of Double array"
    Dim i As Long
    
    MinDouble = arr(0)
    For i = 1 To UBound(arr)
        If MinDouble > arr(i) Then MinDouble = arr(i)
    Next i
End Function

'Function to find a double value in double array (returns a Long : -1 if not found, position in array if found)
Public Function FindDouble(ByRef arr() As Double, ByVal val As Double) As Long
Attribute FindDouble.VB_Description = "Function to find a double value in double array (returns a Long : -1 if not found, position in array if found)"
    Dim i As Long
    
    FindDouble = -1
    For i = 0 To UBound(arr)
        If arr(i) = val Then
            FindDouble = i
            Exit Function
        End If
    Next i
End Function
