Attribute VB_Name = "LongFunctions"
Attribute VB_Description = "Pre-made functions and subs using Long type\n"
Option Explicit

'Function to calculate average value of long array (!CARERUL! returns result as double)
'Works with the "sumLongArray" from LongFunctions
Public Function AverageLong(arr() As Long) As Double
    AverageInt = CDbl(SumIntArray(arr)) / CDbl(UBound(arr))
End Function

'Function to find a Long value in Long array (returns a Long : -1 if not found, position in array if found)
Public Function FindLong(ByRef arr() As Long, ByVal val As Long) As Long
    Dim i As Long

    FindLong = -1
    For i = 0 To UBound(arr)
        If arr(i) = val Then
            FindLong = i
            Exit Function
        End If
    Next i
End Function

'Recursive quicksort for ascending sort of long int array
Public Sub LongQuickSortAsc(ByRef arr() As Long, ByVal leftIndex As Long, ByVal rightIndex As Long)
    Dim i As Long, j As Long, tmp As Long, pivot As Long

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

    If leftIndex < j Then Call LongQuickSortAsc(arr(), leftIndex, j)
    If i < rightIndex Then Call LongQuickSortAsc(arr(), i, rightIndex)
End Sub

'Recursive quicksort for descending sort of long int array
Public Sub LongQuickSortDesc(ByRef arr() As Long, ByVal leftIndex As Long, ByVal rightIndex As Long)
    Dim pivot As Long, tmp As Long, i As Long, j As Long

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

    If (leftIndex < j) Then LongQuickSortDesc arr, leftIndex, j
    If (i < rightIndex) Then LongQuickSortDesc arr, i, rightIndex
End Sub

'Function to generate an long integer (Long) between 2 values (min, max)
Public Function LongRndBetween(ByVal min As Long, ByVal max As Long) As Long
    Randomize
    LongRndBetween = CLng((max - min + 1) * Rnd + min)
End Function

'Sub to sort an array of long integer ascending
'Needs the "longQuickSortAsc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub LongSortAsc(ByRef arr() As Long)
    Call LongQuickSortAsc(arr, 0, UBound(arr))
End Sub

'Sub to sort an array of long integer descending
'Needs the "longQuickSortDesc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub LongSortDesc(ByRef arr() As Long)
    Call LongQuickSortDesc(arr, 0, UBound(arr))
End Sub

'Function to return the max of Long array
Public Function MaxLong(ByRef arr() As Long) As Long
    Dim i As Long

    MaxLong = arr(0)
    For i = 1 To UBound(arr)
        If MaxLong < arr(i) Then MaxLong = arr(i)
    Next i
End Function

'Function to return the min of Long array
Public Function MinLong(ByRef arr() As Long) As Long
    Dim i As Long

    MinLong = arr(0)
    For i = 1 To UBound(arr)
        If MinLong > arr(i) Then MinLong = arr(i)
    Next i
End Function

'Function to sum all values in long array
Public Function SumLongArray(ByRef arr() As Long) As Long
    Dim i As Long

    SumLongArray = 0
    For i = 0 To UBound(arr)
        SumLongArray = SumLongArray + arr(i)
    Next i
End Function
