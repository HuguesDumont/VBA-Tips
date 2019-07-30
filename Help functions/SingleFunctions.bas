Attribute VB_Name = "SingleFunctions"
Attribute VB_Description = "Pre-made functions and subs using single type"
Option Explicit

'Function to generate a single between 2 values (min, max)
Public Function SingleRndBetween(ByVal min As Single, ByVal max As Single) As Single
Attribute SingleRndBetween.VB_Description = "Function to generate a single between 2 values (min, max)"
    Dim rndVariable As Single
    
    Randomize
    rndVariable = Rnd
    
    If (max - min + 1) * rndVariable + min <= max Then
        SingleRndBetween = (max - min + 1) * rndVariable + min
    Else
        Do While (max - min + 1) * rndVariable + min > max
            rndVariable = Rnd
            If (max - min + 1) * rndVariable + min <= max Then
                SingleRndBetween = (max - min + 1) * rndVariable + min
            End If
        Loop
    End If
End Function

'Sub to sort an array of single ascending
'Needs the "singleQuickSortAsc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub SingleSortAsc(ByRef arr() As Single)
Attribute SingleSortAsc.VB_Description = "Sub to sort an array of single ascending\r\nNeeds the ""singleQuickSortAsc"" sub to work\r\nDefault ubound for array is 0 and max index is ubound-1"
    Call SingleQuickSortAsc(arr, 0, UBound(arr))
End Sub

'Sub to sort an array of single descending
'Needs the "singleQuickSortDesc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub SingleSortDesc(ByRef arr() As Single)
Attribute SingleSortDesc.VB_Description = "Sub to sort an array of single descending\r\nNeeds the ""singleQuickSortDesc"" sub to work\r\nDefault ubound for array is 0 and max index is ubound-1"
    Call SingleQuickSortDesc(arr, 0, UBound(arr))
End Sub

'Recursive quicksort for ascending sort of single array
Public Sub SingleQuickSortAsc(ByRef arr() As Single, ByVal leftIndex As Single, ByVal rightIndex As Single)
Attribute SingleQuickSortAsc.VB_Description = "Recursive quicksort for ascending sort of single array"
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
        
    If leftIndex < j Then Call SingleQuickSortAsc(arr(), leftIndex, j)
    If i < rightIndex Then Call SingleQuickSortAsc(arr(), i, rightIndex)
End Sub

'Recursive quicksort for descending sort of single array
Public Sub SingleQuickSortDesc(ByRef arr() As Single, ByVal leftIndex As Single, ByVal rightIndex As Single)
Attribute SingleQuickSortDesc.VB_Description = "Recursive quicksort for descending sort of single array"
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

    If (leftIndex < j) Then SingleQuickSortDesc arr, leftIndex, j
    If (i < rightIndex) Then SingleQuickSortDesc arr, i, rightIndex
End Sub

'Function to sum all values in single array
Public Function SumSingleArray(ByRef arr() As Single) As Single
Attribute SumSingleArray.VB_Description = "Function to sum all values in single array"
    Dim i As Long
    
    SumSingleArray = 0
    For i = 0 To UBound(arr)
        SumSingleArray = SumSingleArray + arr(i)
    Next i
End Function

'Function to calculate average value of single array (!CARERUL! returns result as double)
'Works with the "sumSingleArray" from SingleFunctions
Public Function AverageLong(ByRef arr() As Single) As Double
Attribute AverageLong.VB_Description = "Function to calculate average value of single array (!CARERUL! returns result as double)\r\nWorks with the ""sumSingleArray"" from SingleFunctions"
    AverageInt = CDbl(SumIntArray(arr)) / CDbl(UBound(arr))
End Function

'Function to return the max of single array
Public Function maxSingle(ByRef arr() As Single) As Single
Attribute maxSingle.VB_Description = "Function to return the max of single array"
    Dim i As Long
    
    maxSingle = arr(0)
    For i = 1 To UBound(arr)
        If maxSingle < arr(i) Then maxSingle = arr(i)
    Next i
End Function

'Function to return the min of single array
Public Function MinSingle(ByRef arr() As Single) As Single
Attribute MinSingle.VB_Description = "Function to return the min of single array"
    Dim i As Long
    
    MinSingle = arr(0)
    For i = 1 To UBound(arr)
        If MinSingle > arr(i) Then MinSingle = arr(i)
    Next i
End Function

'Function to find a Single value in Single array (returns a Long : -1 if not found, position in array if found)
Public Function FindSingle(ByRef arr() As Single, ByVal val As Single) As Long
Attribute FindSingle.VB_Description = "Function to find a Single value in Single array (returns a Long : -1 if not found, position in array if found)"
    Dim i As Long
    
    FindSingle = -1
    For i = 0 To UBound(arr)
        If arr(i) = val Then
            FindSingle = i
            Exit Function
        End If
    Next i
End Function
