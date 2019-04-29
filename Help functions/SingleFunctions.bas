Attribute VB_Name = "SingleFunctions"
Attribute VB_Description = "Pre-made functions and subs using single type"
Option Explicit

'Function to generate a single between 2 values (min, max)
Public Function singleRndBetween(min As Single, max As Single) As Single
Attribute singleRndBetween.VB_Description = "Function to generate a single between 2 values (min, max)"
    Dim rndVariable As Single
    
    Randomize
    rndVariable = Rnd
    
    If (max - min + 1) * rndVariable + min <= max Then
        singleRndBetween = (max - min + 1) * rndVariable + min
    Else
        Do While (max - min + 1) * rndVariable + min > max
            rndVariable = Rnd
            If (max - min + 1) * rndVariable + min <= max Then
                singleRndBetween = (max - min + 1) * rndVariable + min
            End If
        Loop
    End If
End Function

'Sub to sort an array of single ascending
'Needs the "singleQuickSortAsc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub singleSortAsc(ByRef arr() As Single)
Attribute singleSortAsc.VB_Description = "Sub to sort an array of single ascending\r\nNeeds the ""singleQuickSortAsc"" sub to work\r\nDefault ubound for array is 0 and max index is ubound-1"
    Call singleQuickSortAsc(arr, 0, UBound(arr))
End Sub

'Sub to sort an array of single descending
'Needs the "singleQuickSortDesc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub singleSortDesc(ByRef arr() As Single)
Attribute singleSortDesc.VB_Description = "Sub to sort an array of single descending\r\nNeeds the ""singleQuickSortDesc"" sub to work\r\nDefault ubound for array is 0 and max index is ubound-1"
    Call singleQuickSortDesc(arr, 0, UBound(arr))
End Sub

'Recursive quicksort for ascending sort of single array
Public Sub singleQuickSortAsc(ByRef arr() As Single, ByVal leftIndex As Single, ByVal rightIndex As Single)
Attribute singleQuickSortAsc.VB_Description = "Recursive quicksort for ascending sort of single array"
    Dim i As Single, j As Single, tmp As Single, pivot As Single
    
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
        
    If leftIndex < j Then Call singleQuickSortAsc(arr(), leftIndex, j)
    If i < rightIndex Then Call singleQuickSortAsc(arr(), i, rightIndex)
End Sub

'Recursive quicksort for descending sort of single array
Public Sub singleQuickSortDesc(arr() As Single, leftIndex As Single, rightIndex As Single)
Attribute singleQuickSortDesc.VB_Description = "Recursive quicksort for descending sort of single array"
    Dim pivot As Single, tmp As Single, i As Single, j As Single

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

    If (leftIndex < j) Then singleQuickSortDesc arr, leftIndex, j
    If (i < rightIndex) Then singleQuickSortDesc arr, i, rightIndex
End Sub

'Function to sum all values in single array
Public Function sumSingleArray(arr() As Single) As Single
Attribute sumSingleArray.VB_Description = "Function to sum all values in single array"
    Dim i As Long
    
    sumSingleArray = 0
    For i = 0 To UBound(arr)
        sumSingleArray = sumSingleArray + arr(i)
    Next i
End Function

'Function to calculate average value of single array (!CARERUL! returns result as double)
'Works with the "sumSingleArray" from SingleFunctions
Public Function averageLong(arr() As Single) As Double
Attribute averageLong.VB_Description = "Function to calculate average value of single array (!CARERUL! returns result as double)\r\nWorks with the ""sumSingleArray"" from SingleFunctions"
    averageInt = CDbl(sumIntArray(arr)) / CDbl(UBound(arr))
End Function

'Function to return the max of single array
Public Function maxSingle(arr() As Single) As Single
Attribute maxSingle.VB_Description = "Function to return the max of single array"
    Dim i As Long
    
    maxSingle = arr(0)
    For i = 1 To UBound(arr)
        If maxSingle < arr(i) Then maxSingle = arr(i)
    Next i
End Function

'Function to return the min of single array
Public Function minSingle(arr() As Single) As Single
Attribute minSingle.VB_Description = "Function to return the min of single array"
    Dim i As Long
    
    minSingle = arr(0)
    For i = 1 To UBound(arr)
        If minSingle > arr(i) Then minSingle = arr(i)
    Next i
End Function

'Function to find a Single value in Single array (returns a Long : -1 if not found, position in array if found)
Public Function findSingle(arr() As Single, val As Single) As Long
Attribute findSingle.VB_Description = "Function to find a Single value in Single array (returns a Long : -1 if not found, position in array if found)"
    Dim i As Long
    
    findSingle = -1
    For i = 0 To UBound(arr)
        If arr(i) = val Then
            findSingle = i
            Exit Function
        End If
    Next i
End Function
