Attribute VB_Name = "Sorting"
Attribute VB_Description = "Functions to sort array"
Option Explicit

'Bubble sort
Public Sub BubbleSort(ByRef arr As Variant)
    Dim changed                         As Boolean
    Dim i                               As Long
    Dim j                               As Long
    Dim tmp                             As Variant
    
    For j = UBound(arr) To 0 Step -1
        changed = False
        
        For i = 0 To j - 1
            If (arr(i) > arr(i + 1)) Then
                tmp = arr(i)
                arr(i) = arr(i + 1)
                arr(i + 1) = tmp
                changed = True
            End If
        Next i
        
        If (Not changed) Then
            Exit For
        End If
    Next j
End Sub

'Insertion sort
Public Sub InsertionSort(ByRef arr As Variant)
    Dim i                               As Long
    Dim j                               As Long
    Dim Temp                            As Variant
    
    For i = 0 To UBound(arr)
        Temp = arr(i)
        j = i - 1
        
        Do While (arr(j) > Temp)
            arr(j + 1) = arr(j)
            j = j - 1
            
            If (j < 0) Then
                Exit Do
            End If
        Loop
        
        arr(j + 1) = Temp
    Next i
End Sub

'Quick sort
Public Sub QuickSort(ByRef arr As Variant)
    Call QuickSortRecursive(arr, 0, UBound(arr))
End Sub

'Recursive quick sort
Private Sub QuickSortRecursive(ByRef arr As Variant, ByVal leftIndex As Variant, ByVal rightIndex As Variant)
    Dim i                               As Variant
    Dim j                               As Variant
    Dim tmp                             As Variant
    Dim pivot                           As Variant
    
    i = leftIndex
    j = rightIndex
    pivot = arr((leftIndex + rightIndex) / 2)
    
    Do
        While (pivot > arr(i))
            i = i + 1
        Wend
        
        While (arr(j) > pivot)
            j = j - 1
        Wend
        
        If ((j + 1) > i) Then
            tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            j = j - 1
            i = i + 1
        End If
    Loop Until (i > j)
    
    If (leftIndex < j) Then Call QuickSortRecursive(arr, leftIndex, j)
    If (i < rightIndex) Then Call QuickSortRecursive(arr, i, rightIndex)
End Sub

'Selection sort
Public Sub SelectionSort(ByRef arr As Variant)
    Dim i                               As Long
    Dim j                               As Long
    Dim tmp                             As Variant
    
    For i = 0 To UBound(arr)
        For j = i + 1 To UBound(arr)
            If (arr(j) < arr(i)) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
End Sub
