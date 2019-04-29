Attribute VB_Name = "Sorting"
Attribute VB_Description = "Functions to sort array"
Option Explicit

'Bubble sort
Public Function bubbleSort(arr As Variant) As Variant
Attribute bubbleSort.VB_Description = "Simple bubble sort"
    Dim i As Long, j As Long
    Dim tmp As Variant
    Dim changed As Boolean
    
    For j = UBound(arr) To 0 Step -1
        changed = False
        
        For i = 0 To j - 1
            If arr(i) > arr(i + 1) Then
                tmp = arr(i)
                arr(i) = arr(i + 1)
                arr(i + 1) = tmp
                changed = True
            End If
        Next i
        
        If Not changed Then
            Exit For
        End If
        
    Next j
    
    bubbleSort = arr
End Function

'Insertion sort
Public Function insertionSort(arr As Variant) As Variant
Attribute insertionSort.VB_Description = "Simple insertion sort"
    Dim Temp As Variant
    Dim i As Long, j As Long
    
    For i = 1 To UBound(arr)
        Temp = arr(i)
        j = i - 1
        Do While arr(j) > Temp
            arr(j + 1) = arr(j)
            j = j - 1
            If j < 0 Then
                Exit Do
            End If
        Loop
        arr(j + 1) = Temp
    Next i
    insertionSort = arr
End Function

'Selection sort
Public Function selectionSort(arr As Variant) As Variant
Attribute selectionSort.VB_Description = "Simple selection sort"
    Dim i As Long, j As Long, k As Long
    Dim tmp As Variant
    
    For i = 0 To UBound(arr)
        k = i
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next
    Next
    selectionSort = arr
End Function

'Quick sort
Public Sub quickSort(ByRef arr As Variant)
Attribute quickSort.VB_Description = "Quick sort using recursiveQuickSort"
    Call quickSortRecursive(arr, 0, UBound(arr))
End Sub

'Recursive quick sort
Private Sub quickSortRecursive(ByRef arr As Variant, ByVal leftIndex As Variant, ByVal rightIndex As Variant)
    Dim i As Variant, j As Variant, tmp As Variant, pivot As Variant
    
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
        
    If leftIndex < j Then Call quickSortRecursive(arr, leftIndex, j)
    If i < rightIndex Then Call quickSortRecursive(arr, i, rightIndex)
End Sub
