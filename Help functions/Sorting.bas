Attribute VB_Name = "Sorting"
Attribute VB_Description = "Functions to sort array"
Option Explicit

'Bubble sort
Public Sub BubbleSort(ByRef arr As Variant)
Attribute BubbleSort.VB_Description = "Simple bubble sort"
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
End Sub

'Insertion sort
Public Sub InsertionSort(ByRef arr As Variant)
Attribute InsertionSort.VB_Description = "Simple insertion sort"
    Dim Temp As Variant
    Dim i As Long, j As Long
    
    For i = 0 To UBound(arr)
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
End Sub

'Selection sort
Public Sub SelectionSort(ByRef arr As Variant)
Attribute SelectionSort.VB_Description = "Simple selection sort"
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
End Sub

'Quick sort
Public Sub QuickSort(ByRef arr As Variant)
Attribute QuickSort.VB_Description = "Quick sort using recursiveQuickSort"
    Call QuickSortRecursive(arr, 0, UBound(arr))
End Sub

'Recursive quick sort
Private Sub QuickSortRecursive(ByRef arr As Variant, ByVal leftIndex As Variant, ByVal rightIndex As Variant)
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
        
    If leftIndex < j Then Call QuickSortRecursive(arr, leftIndex, j)
    If i < rightIndex Then Call QuickSortRecursive(arr, i, rightIndex)
End Sub
