Attribute VB_Name = "ArraysAndCollections"
Attribute VB_Description = "Help functions for Arrays, Collections, ArrayLists and Dictionaries"
Option Explicit

' Add all elements of an array to another array
' Parameters :
' - arr      : the array where to add elements
' - toAdd    : the array of elements to add
Public Sub AddAllArray(ByRef arr As Variant, ByVal toAdd As Variant)
    Dim i                               As Long
    Dim limit                           As Long
    
    If (UBound(toAdd) > -1) Then
        limit = UBound(arr) + 1
        ReDim Preserve arr(0 To UBound(arr) + UBound(toAdd) + 1) As Variant
        
        For i = limit To UBound(arr)
            arr(i) = toAdd(i - limit)
        Next i
    End If
End Sub

' Add an element to array
' Parameters :
' - arr      : the array where to add the element
' - val      : the element to add
Public Sub AddArray(ByRef arr As Variant, ByVal toAdd As Variant)
    ReDim Preserve arr(0 To UBound(arr) + 1) As Variant
    arr(UBound(arr)) = toAdd
End Sub

' Convert a 2D array to a dictionary structure
' Parameters :
' - arr      : the array to convert
' Returns    : a dictionary
Public Function ArrayToDictionary(ByVal arr As Variant) As Scripting.Dictionary
    Dim i                               As Long
    
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

' Complement of the intersection of 2 arrays
' It uses the union, intersection and difference function [difference(union(a,b),intersection(a,b))]
' Parameters :
' - arrLeft  : first array
' - arrRight : second array
' Returns : an array (variant)
Public Function ComplementIntersection(ByVal arrLeft As Variant, ByVal arrRight As Variant) As Variant
    ComplementIntersection = Difference(Union(arrLeft, arrRight), Intersection(arrLeft, arrRight))
End Function

' Check if array contains a specific value
' Parameters :
' - arr      : the array where to search in
' - value    : the element to look for
' Returns : true if array contains the element, else false
Public Function Contains(ByVal arr As Variant, ByVal value As Variant) As Boolean
    Dim i                               As Long
    
    Contains = False
    
    For i = 0 To UBound(arr)
        If (value = arr(i)) Then
            Contains = True
            Exit Function
        End If
    Next i
End Function

' Difference between two arrays
' Parameters :
' - arrLeft  : first array
' - arrRight : second array
' Returns : an array (variant)
Public Function Difference(ByVal arrLeft As Variant, ByVal arrRight As Variant) As Variant
    Dim elem                            As Variant
    
    Difference = arrLeft
    
    For Each elem In arrRight
        If (Contains(Difference, elem)) Then
            Call Remove(Difference, elem)
        End If
    Next elem
End Function

' Get position of an element in array
' Parameters :
' - arr      : the array where to seach in
' - value    : the value to look for
' Returns : the first position of value in array if found, else -1
Public Function IndexOf(ByVal arr As Variant, ByVal value As Variant) As Long
    Dim i                               As Long
    
    IndexOf = -1
    
    For i = 0 To UBound(arr)
        If (value = arr(i)) Then
            IndexOf = i
            Exit Function
        End If
    Next i
End Function

' Intersection of two arrays
' Parameters :
' - arrLeft  : first array
' - arrRight : second array
' Returns : an array (variant)
Public Function Intersection(ByVal arrLeft As Variant, ByVal arrRight As Variant) As Variant
    Dim elem                            As Variant
    
    Intersection = Array()
    
    For Each elem In arrLeft
        If (Contains(arrRight, elem)) Then
            Call AddArray(Intersection, elem)
        End If
    Next elem
End Function

' Remove a specific value from an array (only first instance in the array)
' Parameters :
' - arr      : the array where to search in
' - value    : the element to remove from array
Public Sub Remove(ByRef arr As Variant, ByVal value As Variant)
    Dim i                               As Long
    Dim pos                             As Long
    
    pos = IndexOf(arr, value)
    
    If (pos <> -1) Then
        For i = pos To UBound(arr) - 1
            arr(i) = arr(i + 1)
        Next i
        
        ReDim Preserve arr(UBound(arr) - 1)
    End If
End Sub

' Remove duplicates from an array
' Parameters :
' - arr      : the array from which to remove duplicates
' Returns : an array (variant) without values duplicated
Public Function RemoveDuplicates(ByVal arr As Variant) As Variant
    Dim i                               As Long
    Dim cpt                             As Long
    Dim coll                            As Collection
    Dim tmp                             As Variant
    
    ReDim tmp(UBound(arr))
    
    Set coll = New Collection
    
    For i = LBound(arr) To UBound(arr)
        On Error Resume Next
        coll.Add CStr(arr(i)), CStr(arr(i))
        
        If (Err.number > 0) Then
            On Error GoTo 0
        Else
            tmp(cpt) = arr(i)
            cpt = cpt + 1
        End If
    Next i
    
    ReDim Preserve tmp(cpt - 1)
    RemoveDuplicates = tmp
End Function

' Reverse all elements in array (first element becomes last, last becomes first)
' Parameters :
' - arr      : the array to reverse
' Returns : an array (variant) with its values reversed
Public Function ReverseArray(ByVal arr As Variant) As Variant
    Dim i                               As Long
    Dim j                               As Long
    Dim tmp                             As Variant
    
    j = UBound(arr)
    
    For i = 0 To (j / 2)
        tmp = arr(i)
        arr(i) = arr(j)
        arr(j) = tmp
        j = j - 1
    Next i
    
    ReverseArray = arr
End Function

' Shuffle an array
' Parameters :
' - arr      : the array to shuffle
Public Sub ShuffleArray(ByRef arr As Variant)
    Dim n                               As Long
    Dim pos                             As Long
    Dim tmp                             As Variant
    
    Randomize
    
    For n = 0 To UBound(arr)
        pos = CLng((((UBound(arr)) - n) * Rnd) + n)
        tmp = arr(n)
        arr(n) = arr(pos)
        arr(pos) = tmp
    Next n
End Sub

' Sort a dictionary by keys
' Parameters :
' - dic      : the dictionary to sort
' - order    : the order in which to sort (true/default for ascending sort, false for descending sort)
' Returns : the dictionary sorted
Public Function SortDictionaryByKey(ByVal dic As Object, Optional ByVal sortOrder As Boolean = True) As Scripting.Dictionary
    Dim arrList                         As Object
    Dim key                             As Variant
    
    On Error GoTo errorObjectType
    
    Set arrList = CreateObject("System.Collections.ArrayList")
    
    'Put keys in an ArrayList
    For Each key In dic
        arrList.Add key
    Next key
    
    'Sort the keys
    arrList.Sort
    
    'For descending order, reverse
    If (Not sortOrder) Then
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
    If (Err.number = 450) Then
        Err.Raise 600, "SortDictionaryByKey", "Cannot sort the dictionary if the key is an object"
    End If
End Function

' Sort a dictionary by values
' Parameters :
' - dic      : the dictionary to sort
' - order    : the order in which to sort (true/default for ascending sort, false for descending sort)
' Returns : the dictionary sorted
Public Function SortDictionaryByValue(ByVal dic As Scripting.Dictionary, Optional ByVal sortOrder As Boolean = True) As Scripting.Dictionary
    Dim coll                            As Collection
    Dim tmpDic                          As New Scripting.Dictionary
    Dim arrayList                       As Object
    Dim key                             As Variant
    Dim value                           As Variant
    Dim item                            As Variant
    
    On Error GoTo errorObjectType
    
    Set arrayList = CreateObject("System.Collections.ArrayList")
    
    'Put values in ArrayList and sort
    'Store values in tmpDic with their keys as a collection
    For Each key In dic
        value = dic(key)
        'If the value doesn't exist in dic then add
        If (Not tmpDic.exists(value)) Then
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
    If (Not sortOrder) Then
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
    If (Err.number = 450) Then
        Err.Raise 601, "SortDictionaryByValue", "Cannot sort the dictionary if the value is an object"
    End If
End Function

' Get the union of two arrays
' Parameters :
' - arrLeft  : first array
' - arrRight : second array
' Returns : an array (variant)
Public Function UnionArray(ByVal arrLeft As Variant, ByVal arrRight As Variant) As Variant
    Dim elem                            As Variant
    
    UnionArray = arrLeft
    
    For Each elem In arrRight
        If (Not Contains(UnionArray, elem)) Then
            Call AddArray(UnionArray, elem)
        End If
    Next elem
End Function
