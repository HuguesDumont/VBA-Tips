Attribute VB_Name = "ArraysAndCollections"
Option Explicit

'Function to convert a 2D array to a dictionary
Function ArrayToDictionary(arr() As String) As Scripting.Dictionary
    Dim i As Integer
    
    Set ArrayToDictionary = New Scripting.Dictionary
    
    For i = 0 To UBound(arr) - 1
        ArrayToDictionary.Add arr(i, 0), arr(i, 1)
    Next i
End Function

'Function to sort a dictionary by keys
'0 and default for ascending sort
'1 for descending sort
Function SortDictionaryByKey(dic As Object, Optional order As Integer = 0) As Scripting.Dictionary
    Dim arrList As Object
    Dim key As Variant
    Dim coll As New Collection
    
    Set arrList = CreateObject("System.Collections.ArrayList")
    
    'Put keys in an ArrayList
    For Each key In dic
        arrList.Add key
    Next key
    
    'Sort the keys
    arrList.Sort
    
    'For descending order, reverse
    If sortorder = 1 Then
        arrList.Reverse
    End If
    
    'Init dictionary return
    Set SortDictionaryByKey = New Scripting.Dictionary
    
    'Read through the sorted keys and add to new dictionary
    For Each key In arrList
        SortDictionaryByKey.Add key, dic(key)
    Next key
    
    Set arrList = Nothing
End Function

'Function to sort dictionary by values
'0 and default for ascending sort
'1 for descending sort
Function SortDictionaryByValue(dic As Scripting.Dictionary, Optional sortorder As Integer = 0) As Scripting.Dictionary
    Dim key As Variant, value As Variant
    Dim coll As Collection
    Dim arrayList As Object
    Dim dicTemp As New Scripting.Dictionary
    
    On Error GoTo errorObject
    
    Set arrayList = CreateObject("System.Collections.ArrayList")
   
    'Put values in ArrayList and sort
    'Store values in tempDic with their keys as a collection
    For Each key In dic
        value = dic(key)
        ' if the value doesn't exist in dic then add
        If Not dicTemp.Exists(value) Then
            ' create collection to hold keys, needed for duplicate values
            Set coll = New Collection
            dicTemp.Add value, coll
            arrayList.Add value
        End If
        ' Add the current key to the collection
        dicTemp(value).Add key
    Next key
    
    ' Sort the value
    arrayList.Sort
    
    ' Reverse if descending
    If sortorder = 1 Then
        arrayList.Reverse
    End If
    
    ' Read through the ArrayList and add the values and corresponding
    ' keys from the dicTemp
    Dim item As Variant
    For Each value In arrayList
        Set coll = dicTemp(value)
        For Each item In coll
            SortDictionaryByValue.Add item, value
        Next item
    Next value
    
    Set arrayList = Nothing
    
    Exit Function
    
errorObject:
    If Err.Number = 450 Then
        Err.Raise vbObjectError + 100, "SortDictionaryByValue", "Cannot sort the dictionary if the value is an object"
    End If
End Function

