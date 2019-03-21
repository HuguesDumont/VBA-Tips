Attribute VB_Name = "StringFunctions"
Option Explicit

'Function to generate random char (optional string parameter to select a random char in this string)
'Needs the "intRndBetween" function in "IntFunctions" to work
'The different choices are :
'  - 0 : any printable ascii char (including extended ascii table)
'  - 1 : any printable ascii char (including extended ascii table, excluding blank char like spaces)
'  - 2 : any printable ascii char (excluding extended ascii table)
'  - 3 : any printable ascii char (excluding ascii table, excluding blank char like spaces)
'  - 4 : any ascii alphabetical or numeric char (including extended ascii characters like diacritics, lower and upper char included)
'  - 5 : any ascii alphabetical or numeric char (excluding extended ascii table, including lower and upper char)
'  - 6 : any ascii upper alphabetical or numeric char (including extended ascii table)
'  - 7 : any ascii lower alphabetical or numeric char (including extended ascii table)
'  - 8 : any ascii upper alphabetical or numeric char (excluding extended ascii table)
'  - 9 : any ascii lower alphabetical or numeric char (excluding extended ascii table)
'  - 10 : any ascii alphabetical char (including extended ascii table)
'  - 11 : any ascii upper alphabetical char (including extended ascii table)
'  - 12 : any ascii lower alphabetical char (including extended ascii table)
'  - 13 : any ascii alphabetical char (excluding extended ascii table)
'  - 14 : any ascii upper alphabetical char (excluding extended ascii table)
'  - 15 : any ascii lower alphabetical char (excluding extended ascii table)
'  - 16 : any numeric char (from 0 to 9)
'  - 17 : any special printable char (including extended ascii table ==> including diacritics, excluding standard numerics, excluding blank char like spaces, tabs or cariage returns)
Public Function rndChar(choice As Integer) As String
Attribute rndChar.VB_Description = "This function returns a random char from different string categories. Choices are : \r- 0 : any printable ascii char (including extended ascii table)\r- 1 : any printable ascii char (including extended ascii table, excluding blank char like spaces)"
Attribute rndChar.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim rngs(18) As String
    
    rngs(0) = " !" & Chr(34) & "#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~€‚ƒ„…†‡ˆ‰Š‹Œ‘’“”•–—˜™š›œŸ ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖ×ØÙÚÛÜİŞßàáâãäåæçèéêëìíîïğñòóôõö÷øùúûüışÿ"
    rngs(1) = "!" & Chr(34) & "#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~€‚ƒ„…†‡ˆ‰Š‹Œ‘’“”•–—˜™š›œŸ ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖ×ØÙÚÛÜİŞßàáâãäåæçèéêëìíîïğñòóôõö÷øùúûüışÿ"
    rngs(2) = " !" & Chr(34) & "#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
    rngs(3) = "!" & Chr(34) & "#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
    rngs(4) = "abcdefghijklmnopqrstuvwxyzœšÿàáâãäåæçèéêëìíîïğñòóôõöøùúûüışABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789ŒŠŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖØÙÚÛÜİŞ"
    rngs(5) = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    rngs(6) = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789ŒŠŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖØÙÚÛÜİŞ"
    rngs(7) = "abcdefghijklmnopqrstuvwxyz0123456789œšÿàáâãäåæçèéêëìíîïğñòóôõöøùúûüış"
    rngs(8) = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    rngs(9) = "abcdefghijklmnopqrstuvwxyz0123456789"
    rngs(10) = "abcdefghijklmnopqrstuvwxyzœšÿàáâãäåæçèéêëìíîïğñòóôõöøùúûüışABCDEFGHIJKLMNOPQRSTUVWXYZŒŠŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖØÙÚÛÜİŞ"
    rngs(11) = "ABCDEFGHIJKLMNOPQRSTUVWXYZŒŠŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖØÙÚÛÜİŞ"
    rngs(12) = "abcdefghijklmnopqrstuvwxyzœšÿàáâãäåæçèéêëìíîïğñòóôõöøùúûüış"
    rngs(13) = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    rngs(14) = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    rngs(15) = "abcdefghijklmnopqrstuvwxyz"
    rngs(16) = "0123456789"
    rngs(17) = "!" & Chr(34) & "#$%&'()*+,-./:;<=>?@[\]^_`{|}~€‚ƒ„…†‡ˆ‰Š‹Œ‘’“”•–—˜™š›œŸ ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖ×ØÙÚÛÜİŞßàáâãäåæçèéêëìíîïğñòóôõö÷øùúûüışÿ"

    rndChar = Mid(rngs(choice), intFunctions.intRndBetween(1, Len(rngs(choice))), 1) 'Using the "intRndBetween" from "IntFunctions" here
End Function

'sub to sort an array of string in lexicogrpahic order (using extended ascii table values of char in ascending order)
'Needs the "stringQuickSortAsc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub stringSortAsc(ByRef arr() As String)
    Call stringQuickSortAsc(arr, 0, UBound(arr) - 1)
End Sub

'sub to sort an array of string in reverse lexicographic order (using extended ascii table values of char in descending order)
'Needs the "stringQuickSortDesc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub stringSortDesc(ByRef arr() As String)
    Call stringQuickSortDesc(arr, 0, UBound(arr) - 1)
End Sub

'recursive quicksort for lexicographic sort of strig array (using extended ascii table values of char in ascending order)
Public Sub stringQuickSortAsc(ByRef arr() As String, ByVal leftIndex As Long, ByVal rightIndex As Long)
    Dim i As Long, j As Long, tmp As String, pivot As String
    
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
        
    If leftIndex < j Then Call stringQuickSortAsc(arr(), leftIndex, j)
    If i < rightIndex Then Call stringQuickSortAsc(arr(), i, rightIndex)
End Sub

'recursive quicksort for reverse lexicographic sort of strig array (using extended ascii table values of char in descending order)
Public Sub stringQuickSortDesc(arr() As String, leftIndex As Long, rightIndex As Long)
    Dim pivot As String, tmp As String, i As Long, j As Long

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

    If (leftIndex < j) Then stringQuickSortDesc arr, leftIndex, j
    If (i < rightIndex) Then stringQuickSortDesc arr, i, rightIndex
End Sub

'More powerful sort of string is to use the System.Collections.ArrayList (lexicographical ascending sort using ascii table)
'Example of using the System.Collections.ArrayList :
Public Sub exampleArrayList()
    Dim arrList As Object
    Dim i As Integer
    
    Set arrList = CreateObject("System.Collections.ArrayList")
    
    arrList.Add "reogbeor"
    arrList.Add "onreogn"
    arrList.Add "354"
    arrList.Add "1520"
    arrList.Add "anfeomz"
    arrList.Add "à_çu_"
    arrList.Add "!ren:rze,re"

    arrList.Sort
    
    For i = 0 To arrList.Count - 1
        Debug.Print arrList.Item(i)
    Next i
End Sub

'Function to generate random password (not perfect, but works fine and generate really hard passwords to crack)
'If no argument is passed then the function generates a random password composed of 8 char with at least 1 upper char, 1 lower char, 1 special char and 1 numeric
'About args :
'  - "taille" ==> exact length of the password you want
'  - "lowC" ==> mininum number of small letters you want in your password
'  - "upC" ==> minimum number of capital letters you want in your password
'  - "num" ==> minimum number of numeric characters you want in your password
'  - "spec" ==> minimum number of special characters you want in your password
'Works with the "rndChar" function from "StringFunctions" and the "shuffleString" from "StringFunctions"
Public Function randomPass(Optional ByVal taille As Integer = 8, Optional ByVal lowC As Integer = 1, Optional ByVal upC As Integer = 1, _
                            Optional ByVal num As Integer = 1, Optional ByVal spec As Integer = 1) As String
    Dim i As Long, minChar As Long
    
    randomPass = ""
    
    minChar = lowC + upC + num + spec
    'Verifying that wanted length can get all char conditions
    If taille < minChar Then
        MsgBox "Wanted password length is " & taille & ", but conditions on characters are higher (total = " & minChar & ")." & _
            "Password can't be generated. Returned string is empty (" & Chr(34) & Chr(34) & ").", _
            vbCritical + vbOKOnly, "Password couldn't be generated !"
        Exit Function
    End If
    
    'Generating minimum small letters
    While lowC > 0
        randomPass = randomPass & StringFunctions.rndChar(15)
        lowC = lowC - 1
    Wend
    
    'Generating minimum capital letters
    While upC > 0
        randomPass = randomPass & StringFunctions.rndChar(14)
        upC = upC - 1
    Wend
    
    'Generating minimum numeric characters
    While num > 0
        randomPass = randomPass & StringFunctions.rndChar(16)
        num = num - 1
    Wend
    
    'Generating minimum special characters
    While spec > 0
        randomPass = randomPass & StringFunctions.rndChar(17)
        spec = spec - 1
    Wend
    
    While Len(randomPass) <= taille
        randomPass = randomPass & StringFunctions.rndChar(1)
    Wend
    randomPass = shuffleString(randomPass)
End Function

Sub test()
    Debug.Print StringFunctions.randomPass
End Sub

'Function to shuffle a string randomly
'Works with the "LongRndBetween" from "LongFunctions", the "stringToArray" function from "StringFunctions"
'          and the "arrayToString" from "StringFunctions"
Public Function shuffleString(val As String) As String
    Dim n As Long, j As Long
    Dim tmp As String
    Dim tmpArr() As String

    tmpArr = StringFunctions.stringToArray(val)
    
    For n = 0 To UBound(tmpArr) - 1
        j = LongFunctions.longRndBetween(n, UBound(tmpArr) - 1)
        If n <> j Then
            tmp = tmpArr(n)
            tmpArr(n) = tmpArr(j)
            tmpArr(j) = tmp
        End If
    Next n
    
    shuffleString = stringArrayToString(tmpArr)
End Function

'Function to convert string to an array within each char is separated
Public Function stringToArray(val As String) As Variant
    Dim i As Long
    Dim tmp() As String
    
    ReDim tmp(Len(val) - 1)
    
    For i = 1 To Len(val)
        tmp(i - 1) = Mid(val, i, 1)
    Next i
    
    stringToArray = tmp
End Function

'Function to convert a string array to string
Public Function stringArrayToString(arr() As String) As String
    Dim i As Long
    
    stringArrayToString = ""
    For i = 0 To UBound(arr) - 1
        stringArrayToString = stringArrayToString & CStr(arr(i))
    Next i
End Function

'Function to convert a variant array to string
Public Function variantArrayToString(arr() As Variant) As String
    Dim i As Long
    
    variantArrayToString = ""
    For i = 0 To UBound(arr) - 1
        variantArrayToString = variantArrayToString & CStr(arr(i))
    Next i
End Function

'Function to find a String value in String array (returns a Long : -1 if not found, position in array if found)
Public Function findString(arr() As String, val As String) As Long
    Dim i As Long
    
    findString = -1
    For i = 0 To UBound(arr) - 1
        If arr(i) = val Then
            findString = i
            Exit Function
        End If
    Next i
End Function
