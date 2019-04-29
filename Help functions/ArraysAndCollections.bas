Attribute VB_Name = "ArraysAndCollections"
Attribute VB_Description = "Help functions for Arrays, Collections, ArrayLists and Dictionaries"
Option Explicit

'Function to convert a 2D array to a dictionary
Function ArrayToDictionary(arr As Variant) As Scripting.Dictionary
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

'Function to sort a dictionary by keys
'0 and default for ascending sort
'1 for descending sort
Function SortDictionaryByKey(dic As Object, Optional order As Integer = 0) As Scripting.Dictionary
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
Function SortDictionaryByValue(dic As Scripting.Dictionary, Optional order As Integer = 0) As Scripting.Dictionary
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

'Function to remove duplicates from an array
Public Function removeDuplicates(arr As Variant) As Variant
Attribute removeDuplicates.VB_Description = "Function to remove duplicates from an array"
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
    removeDuplicates = tmp
End Function

'Function to reverse array
Public Function reverseArray(arr As Variant) As Variant
Attribute reverseArray.VB_Description = "Function to reverse array"
    Dim i As Long, j As Long
    Dim tmp As Variant
    
    j = UBound(arr)
    For i = 0 To (j / 2)
        tmp = arr(i)
        arr(i) = arr(j)
        arr(j) = tmp
        j = j - 1
    Next i
    
    reverseArray = arr
End Function

'Function to shuffle an array
Public Function shuffleArray(arr As Variant) As Variant
Attribute shuffleArray.VB_Description = "Function to shuffle an array"
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
    
    shuffleArray = tmpArr
End Function

'Function to check if array contains value
Public Function contains(arr As Variant, value As Variant) As Boolean
Attribute contains.VB_Description = "Function to check if array contains value"
    Dim i As Long
    
    contains = False
    For i = 0 To UBound(arr)
        If value = arr(i) Then
            contains = True
            Exit Function
        End If
    Next i
End Function

'Function to get position of first index of a value in an array
'If value is not found in array, return -1
Public Function indexOf(arr As Variant, value As Variant) As Long
Attribute indexOf.VB_Description = "Function to get position of first index of a value in an array\r\nIf value is not found in array, return -1"
    Dim i As Long
    
    indexOf = -1
    For i = 0 To UBound(arr)
        If value = arr(i) Then
            indexOf = i
            Exit Function
        End If
    Next i
End Function

'Function to remove a specific value from an array (only first instance in the array)
Public Function remove(arr As Variant, value As Variant) As Variant
Attribute remove.VB_Description = "Function to remove a specific value from an array (only first instance in the array)"
    Dim i As Long, pos As Long
    Dim tmpArr As Variant
    
    pos = indexOf(arr, value)
    
    If pos = -1 Then
        ReDim tmpArr(UBound(arr))
        For i = 0 To UBound(arr) - 1
            tmpArr(i) = arr(i)
        Next i
    Else
        ReDim tmpArr(UBound(arr) - 1)
        For i = 0 To pos - 1
            tmpArr(i) = arr(i)
        Next i
        For i = pos + 1 To UBound(arr) - 1
            tmpArr(i) = arr(i)
        Next i
    End If
    
    remove = tmpArr
End Function
