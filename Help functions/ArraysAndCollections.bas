Attribute VB_Name = "ArraysAndCollections"
Attribute VB_Description = "Help functions for Arrays, Collections, ArrayLists and Dictionaries"
Option Explicit

'Function to add all elements to array
Public Sub AddAllArray(ByVal toAdd As Variant, ByRef arr As Variant)
Attribute AddAllArray.VB_Description = "Function to add all elements to array"
    Dim i As Long
    Dim limit As Long
    
    limit = UBound(arr) + 1
    
    ReDim Preserve arr(1 To UBound(arr) + UBound(toAdd) + 2) As Variant
    
    For i = limit + 1 To UBound(arr)
        arr(i) = toAdd(i - limit - 1)
    Next i
End Sub

'Function to add an element to array
Public Sub AddArray(ByVal val As Variant, ByRef arr As Variant)
Attribute AddArray.VB_Description = "Function to add an element to array"
    ReDim Preserve arr(0 To UBound(arr) + 1) As Variant
    arr(UBound(arr)) = val
End Sub

'Function to convert a 2D array to a dictionary
Public Function ArrayToDictionary(ByVal arr As Variant) As Scripting.Dictionary
Attribute ArrayToDictionary.VB_Description = "Function to convert a 2D array to a dictionary"
    Dim i As Integer
    
    On Error GoTo arrayDimensionError
    Set ArrayToDictionary = New Scripting.Dictionary
    
    For i = 0 To UBound(arr)
        ArrayToDictionary.Add arr(i, 0), arr(i, 1)
    Next i
    
    Exit Function
    
arrayDimensionError:
    MsgBox "Error while converting array to dictionary. Please verify that you're passing a 2D array." & Chr(13) & _
        "If you keep getting an error, please check on your admin", vbCritical + vbOKOnly, "Error converting array to dictionary"
End Function

'Function to get the complement of the intersection of 2 arrays (union-intersection)
'It uses the union, intersection and difference function [difference(union(a,b),intersection(a,b))]
Public Function ComplementIntersection(ByVal arrLeft As Variant, ByVal arrRight As Variant) As Variant
Attribute ComplementIntersection.VB_Description = "Function to get the complement of the intersection of 2 arrays (union-intersection)\r\nIt uses the union, intersection and difference function [difference(union(a,b),intersection(a,b))]"
    ComplementIntersection = Difference(Union(arrLeft, arrRight), Intersection(arrLeft, arrRight))
End Function

'Function to check if array contains value
Public Function Contains(ByVal arr As Variant, ByVal value As Variant) As Boolean
Attribute Contains.VB_Description = "Function to check if array contains value"
    Dim i As Long
    
    Contains = False
    For i = 0 To UBound(arr)
        If value = arr(i) Then
            Contains = True
            Exit Function
        End If
    Next i
End Function

'Function to get the difference between arrLeft and arrRight
Public Function Difference(ByVal arrLeft As Variant, ByVal arrRight As Variant) As Variant
Attribute Difference.VB_Description = "Function to get the difference between arrLeft and arrRight"
    Dim elem As Variant
    
    Difference = arrLeft
    For Each elem In arrRight
        If (Contains(Difference, elem)) Then
            Call Remove(Difference, elem)
        End If
    Next elem
End Function

'Function to get position of first index of a value in an array
'If value is not found in array, return -1
Public Function IndexOf(ByVal arr As Variant, ByVal value As Variant) As Long
Attribute IndexOf.VB_Description = "Function to get position of first index of a value in an array\r\nIf value is not found in array, return -1"
    Dim i As Long
    
    IndexOf = -1
    For i = 0 To UBound(arr)
        If value = arr(i) Then
            IndexOf = i
            Exit Function
        End If
    Next i
End Function

'Function to get the intersection of two arrays
Public Function Intersection(ByVal arrLeft As Variant, ByVal arrRight As Variant) As Variant
Attribute Intersection.VB_Description = "Function to get the intersection of two arrays"
    Dim elem As Variant
    
    Intersection = Array()
    For Each elem In arrLeft
        If (Contains(arrRight, elem)) Then
            Call AddArray(elem, Intersection)
        End If
    Next elem
End Function

'Function to remove a specific value from an array (only first instance in the array)
Public Sub Remove(ByRef arr As Variant, ByVal value As Variant)
Attribute Remove.VB_Description = "Function to remove a specific value from an array (only first instance in the array)"
    Dim i As Long, pos As Long
    Dim tmpArr As Variant
    
    pos = IndexOf(arr, value)
    
    If (pos <> -1) Then
        For i = pos To UBound(arr) - 1
            arr(i) = arr(i + 1)
        Next i
        ReDim Preserve arr(UBound(arr) - 1)
    End If
End Sub

'Function to remove duplicates from an array
Public Function RemoveDuplicates(ByVal arr As Variant) As Variant
Attribute RemoveDuplicates.VB_Description = "Function to remove duplicates from an array"
    Dim coll As New Collection
    Dim i As Long, cpt As Long
    Dim tmp As Variant
    
    ReDim tmp(UBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        On Error Resume Next
        coll.Add CStr(arr(i)), CStr(arr(i))
        
        If Err.number > 0 Then
            On Error GoTo 0
        Else
            tmp(cpt) = arr(i)
            cpt = cpt + 1
        End If
    Next i
    
    ReDim Preserve tmp(cpt - 1)
    RemoveDuplicates = tmp
End Function

'Function to reverse array
Public Function ReverseArray(ByVal arr As Variant) As Variant
Attribute ReverseArray.VB_Description = "Function to reverse array"
    Dim i As Long, j As Long
    Dim tmp As Variant
    
    j = UBound(arr)
    For i = 0 To (j / 2)
        tmp = arr(i)
        arr(i) = arr(j)
        arr(j) = tmp
        j = j - 1
    Next i
    
    ReverseArray = arr
End Function

'Function to shuffle an array
Public Function ShuffleArray(ByVal arr As Variant) As Variant
Attribute ShuffleArray.VB_Description = "Function to shuffle an array"
    Dim n As Long, j As Long
    Dim tmp As Variant, tmpArr As Variant
    
    Randomize
    
    ReDim tmpArr(UBound(arr))
    
    For n = 0 To UBound(arr)
        tmpArr(n) = arr(n)
    Next n
    
    For n = 0 To UBound(arr)
        j = CLng((((UBound(arr)) - n) * Rnd) + n)
        tmp = tmpArr(n)
        tmpArr(n) = tmpArr(j)
        tmpArr(j) = tmp
    Next n
    
    ShuffleArray = tmpArr
End Function

'Function to sort a dictionary by keys
'0 and default for ascending sort
'1 for descending sort
Public Function SortDictionaryByKey(ByVal dic As Object, Optional ByVal order As Integer = 0) As Scripting.Dictionary
Attribute SortDictionaryByKey.VB_Description = "Function to sort a dictionary by keys\r\n0 and default for ascending sort\r\n1 for descending sort"
    Dim arrList As Object
    Dim key As Variant
    Dim coll As New Collection
    
    On Error GoTo errorObjectType
    
    Set arrList = CreateObject("System.Collections.ArrayList")
    
    'Put keys in an ArrayList
    For Each key In dic
        arrList.Add key
    Next key
    
    'Sort the keys
    arrList.Sort
    
    'For descending order, reverse
    If order = 1 Then
        arrList.Reverse
    End If
    
    'Init dictionary return
    Set SortDictionaryByKey = New Scripting.Dictionary
    
    'Read through the sorted keys and add to new dictionary
    For Each key In arrList
        SortDictionaryByKey.Add key, dic(key)
    Next key
    
    Set arrList = Nothing
    
    Exit Function
    
errorObjectType:
    If Err.number = 450 Then
        Err.Raise vbObjectError + 100, "SortDictionaryByValue", "Cannot sort the dictionary if the value is an object"
    End If
End Function

'Function to sort dictionary by values
'0 and default for ascending sort
'1 for descending sort
Public Function SortDictionaryByValue(ByVal dic As Scripting.Dictionary, Optional ByVal order As Integer = 0) As Scripting.Dictionary
Attribute SortDictionaryByValue.VB_Description = "Function to sort dictionary by values\r\n0 and default for ascending sort\r\n1 for descending sort"
    Dim arrayList As Object
    Dim tmpDic As New Scripting.Dictionary
    Dim key As Variant, value As Variant, item As Variant
    Dim coll As Collection
    
    On Error GoTo errorObjectType
    
    Set arrayList = CreateObject("System.Collections.ArrayList")
   
    'Put values in ArrayList and sort
    'Store values in tmpDic with their keys as a collection
    For Each key In dic
        value = dic(key)
        'If the value doesn't exist in dic then add
        If tmpDic.exists(value) = False Then
            'Create collection to hold keys, needed for duplicate values
            Set coll = New Collection
            tmpDic.Add value, coll
            'Add the value
            arrayList.Add value
        End If
        'Add the current key to the collection
        tmpDic(value).Add key
    Next key
    
    'Sort the values
    arrayList.Sort
    
    'Reverse sort if descending
    If order = 1 Then
        arrayList.Reverse
    End If
    
    Set SortDictionaryByValue = New Scripting.Dictionary
    
    'Read through the ArrayList and add the values and corresponding keys from the tmpDic
    For Each value In arrayList
        Set coll = tmpDic(value)
        For Each item In coll
            SortDictionaryByValue.Add item, value
        Next item
    Next value
    
    Set arrayList = Nothing
    
    Exit Function
    
errorObjectType:
    If Err.number = 450 Then
        Err.Raise vbObjectError + 100, "SortDictionaryByValue", "Cannot sort the dictionary if the value is an object"
    End If
End Function

'Function to get the union between two arrays
Public Function Union(ByVal arrLeft As Variant, ByVal arrRight As Variant) As Variant
Attribute Union.VB_Description = "Function to get the union between two arrays"
    Dim elem As Variant
    
    Union = arrLeft
    For Each elem In arrRight
        If (Not Contains(Union, elem)) Then
            AddArray elem, Union
        End If
    Next elem
End Function
