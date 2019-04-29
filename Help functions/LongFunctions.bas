Attribute VB_Name = "LongFunctions"
Attribute VB_Description = "Pre-made functions and subs using Long type\n"
Option Explicit

'Function to generate an long integer (Long) between 2 values (min, max)
Public Function longRndBetween(min As Long, max As Long) As Long
Attribute longRndBetween.VB_Description = "Function to generate an long integer (Long) between 2 values (min, max)"
    Randomize
    longRndBetween = CLng((max - min + 1) * Rnd + min)
End Function

'Sub to sort an array of long integer ascending
'Needs the "longQuickSortAsc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub longSortAsc(ByRef arr() As Long)
Attribute longSortAsc.VB_Description = "Sub to sort an array of long integer ascending\r\nNeeds the ""longQuickSortAsc"" sub to work\r\nDefault ubound for array is 0 and max index is ubound-1"
    Call longQuickSortAsc(arr, 0, UBound(arr))
End Sub

'Sub to sort an array of long integer descending
'Needs the "longQuickSortDesc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub longSortDesc(ByRef arr() As Long)
Attribute longSortDesc.VB_Description = "Sub to sort an array of long integer descending\r\nNeeds the ""longQuickSortDesc"" sub to work\r\nDefault ubound for array is 0 and max index is ubound-1"
    Call longQuickSortDesc(arr, 0, UBound(arr))
End Sub

'Recursive quicksort for ascending sort of long int array
Public Sub longQuickSortAsc(ByRef arr() As Long, ByVal leftIndex As Long, ByVal rightIndex As Long)
Attribute longQuickSortAsc.VB_Description = "Recursive quicksort for ascending sort of long int array"
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
        
    If leftIndex < j Then Call longQuickSortAsc(arr(), leftIndex, j)
    If i < rightIndex Then Call longQuickSortAsc(arr(), i, rightIndex)
End Sub

'Recursive quicksort for descending sort of long int array
Public Sub longQuickSortDesc(arr() As Long, leftIndex As Long, rightIndex As Long)
Attribute longQuickSortDesc.VB_Description = "Recursive quicksort for descending sort of long int array"
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

    If (leftIndex < j) Then longQuickSortDesc arr, leftIndex, j
    If (i < rightIndex) Then longQuickSortDesc arr, i, rightIndex
End Sub

'Function to sum all values in long array
Public Function sumLongArray(arr() As Long) As Long
Attribute sumLongArray.VB_Description = "Function to sum all values in long array"
    Dim i As Long
    
    sumLongArray = 0
    For i = 0 To UBound(arr)
        sumLongArray = sumLongArray + arr(i)
    Next i
End Function

'Function to calculate average value of long array (!CARERUL! returns result as double)
'Works with the "sumLongArray" from LongFunctions
Public Function averageLong(arr() As Long) As Double
Attribute averageLong.VB_Description = "Function to calculate average value of long array (!CARERUL! returns result as double)\r\nWorks with the ""sumLongArray"" from LongFunctions"
    averageInt = CDbl(sumIntArray(arr)) / CDbl(UBound(arr))
End Function

'Function to return the max of Long array
Public Function maxLong(arr() As Long) As Long
Attribute maxLong.VB_Description = "Function to return the max of Long array"
    Dim i As Long
    
    maxLong = arr(0)
    For i = 1 To UBound(arr)
        If maxLong < arr(i) Then maxLong = arr(i)
    Next i
End Function

'Function to return the min of Long array
Public Function minLong(arr() As Long) As Long
Attribute minLong.VB_Description = "Function to return the min of Long array"
    Dim i As Long
    
    minLong = arr(0)
    For i = 1 To UBound(arr)
        If minLong > arr(i) Then minLong = arr(i)
    Next i
End Function

'Function to find a Long value in Long array (returns a Long : -1 if not found, position in array if found)
Public Function findLong(arr() As Long, val As Long) As Long
Attribute findLong.VB_Description = "Function to find a Long value in Long array (returns a Long : -1 if not found, position in array if found)"
    Dim i As Long
    
    findLong = -1
    For i = 0 To UBound(arr)
        If arr(i) = val Then
            findLong = i
            Exit Function
        End If
    Next i
End Function
