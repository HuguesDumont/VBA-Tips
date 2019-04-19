Attribute VB_Name = "ArraysAndCollections"
Attribute VB_Description = "Help functions for Arrays, Collections, ArrayLists and Dictionaries"
Option Explicit

'Function to convert a 2D array to a dictionary
Function ArrayToDictionary(arr() As String) As Scripting.Dictionary
Attribute ArrayToDictionary.VB_Description = "Function to convert a 2D array to a dictionary"
    Dim i As Integer
    
    On Error GoTo arrayDimensionError
    Set ArrayToDictionary = New Scripting.Dictionary
    
    For i = 0 To UBound(arr) - 1
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
    If Err.Number = 450 Then
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
    If Err.Number = 450 Then
        Err.Raise vbObjectError + 100, "SortDictionaryByValue", "Cannot sort the dictionary if the value is an object"
    End If
End Function
