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
Public Function RndChar(ByVal choice As Integer) As String
    Dim rngs(18) As String
    
    rngs(0) = " !" & Chr(34) & "#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}" & _
        "~€‚ƒ„…†‡ˆ‰Š‹Œ‘’“”•–—˜™š›œŸ ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖ×ØÙÚÛÜİŞßàáâãäåæçèéêëìíîïğñòóôõö÷øùúûüışÿ"
    rngs(1) = "!" & Chr(34) & "#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}" & _
        "~€‚ƒ„…†‡ˆ‰Š‹Œ‘’“”•–—˜™š›œŸ ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖ×ØÙÚÛÜİŞßàáâãäåæçèéêëìíîïğñòóôõö÷øùúûüışÿ"
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

    RndChar = Mid(rngs(choice), intFunctions.IntRndBetween(1, Len(rngs(choice))), 1) 'Using the "intRndBetween" from "IntFunctions" here
End Function

'sub to sort an array of string in lexicogrpahic order (using extended ascii table values of char in ascending order)
'Needs the "stringQuickSortAsc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub StringSortAsc(ByRef arr() As String)
    Call StringQuickSortAsc(arr, 0, UBound(arr) - 1)
End Sub

'sub to sort an array of string in reverse lexicographic order (using extended ascii table values of char in descending order)
'Needs the "stringQuickSortDesc" sub to work
'Default ubound for array is 0 and max index is ubound-1
Public Sub StringSortDesc(ByRef arr() As String)
    Call StringQuickSortDesc(arr, 0, UBound(arr) - 1)
End Sub

'recursive quicksort for lexicographic sort of strig array (using extended ascii table values of char in ascending order)
Public Sub StringQuickSortAsc(ByRef arr() As String, ByVal leftIndex As Long, ByVal rightIndex As Long)
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
        
    If leftIndex < j Then Call StringQuickSortAsc(arr(), leftIndex, j)
    If i < rightIndex Then Call StringQuickSortAsc(arr(), i, rightIndex)
End Sub

'recursive quicksort for reverse lexicographic sort of strig array (using extended ascii table values of char in descending order)
Public Sub StringQuickSortDesc(ByRef arr() As String, ByVal leftIndex As Long, ByVal rightIndex As Long)
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

    If (leftIndex < j) Then StringQuickSortDesc arr, leftIndex, j
    If (i < rightIndex) Then StringQuickSortDesc arr, i, rightIndex
End Sub

'Function to generate random password (not perfect, but works fine and generate really hard passwords to crack)
'If no argument is passed then the function generates a random password composed of 8 char with at least 1 upper char, 1 lower char, 1 special char and 1 numeric
'About args :
'- "taille" ==> exact length of the password you want
'- "lowC" ==>   mininum number of small letters you want in your password
'- "upC" ==>    minimum number of capital letters you want in your password
'- "num" ==>    minimum number of numeric characters you want in your password
'- "spec" ==>   minimum number of special characters you want in your password
'Works with the "rndChar" function from "StringFunctions" and the "shuffleString" from "StringFunctions"
Public Function RandomPass(Optional ByVal taille As Integer = 8, Optional ByVal lowC As Integer = 1, Optional ByVal upC As Integer = 1, _
        Optional ByVal num As Integer = 1, Optional ByVal spec As Integer = 1) As String
    Dim i As Long, minChar As Long
    
    RandomPass = ""
    
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
        RandomPass = RandomPass & StringFunctions.RndChar(15)
        lowC = lowC - 1
    Wend
    
    'Generating minimum capital letters
    While upC > 0
        RandomPass = RandomPass & StringFunctions.RndChar(14)
        upC = upC - 1
    Wend
    
    'Generating minimum numeric characters
    While num > 0
        RandomPass = RandomPass & StringFunctions.RndChar(16)
        num = num - 1
    Wend
    
    'Generating minimum special characters
    While spec > 0
        RandomPass = RandomPass & StringFunctions.RndChar(17)
        spec = spec - 1
    Wend
    
    While Len(RandomPass) <= taille
        RandomPass = RandomPass & StringFunctions.RndChar(1)
    Wend
    RandomPass = ShuffleString(RandomPass)
End Function

'Function to shuffle a string randomly
'Works with the "LongRndBetween" from "LongFunctions", the "stringToArray" function from "StringFunctions"
'          and the "arrayToString" from "StringFunctions"
Public Function ShuffleString(ByVal val As String) As String
    Dim n As Long, j As Long
    Dim tmp As String
    Dim tmpArr() As String

    tmpArr = StringFunctions.StringToArray(val)
    
    For n = 0 To UBound(tmpArr) - 1
        j = LongFunctions.LongRndBetween(n, UBound(tmpArr) - 1)
        If n <> j Then
            tmp = tmpArr(n)
            tmpArr(n) = tmpArr(j)
            tmpArr(j) = tmp
        End If
    Next n
    
    ShuffleString = StringArrayToString(tmpArr)
End Function

'Function to convert string to an array within each char is separated
Public Function StringToArray(ByVal val As String) As Variant
    Dim i As Long
    Dim tmp() As String
    
    ReDim tmp(Len(val) - 1)
    
    For i = 1 To Len(val)
        tmp(i - 1) = Mid(val, i, 1)
    Next i
    StringToArray = tmp
End Function

'Function to convert a string array to string
Public Function StringArrayToString(ByRef arr() As String) As String
    Dim i As Long
    
    StringArrayToString = ""
    For i = 0 To UBound(arr) - 1
        StringArrayToString = StringArrayToString & CStr(arr(i))
    Next i
End Function

'Function to convert a variant array to string
Public Function VariantArrayToString(arr() As Variant) As String
    Dim i As Long
    
    VariantArrayToString = ""
    For i = 0 To UBound(arr) - 1
        VariantArrayToString = VariantArrayToString & CStr(arr(i))
    Next i
End Function

'Function to find a String value in String array (returns a Long : -1 if not found, position in array if found)
Public Function FindString(ByRef arr() As String, ByVal val As String) As Long
    Dim i As Long
    
    FindString = -1
    For i = 0 To UBound(arr) - 1
        If arr(i) = val Then
            FindString = i
            Exit Function
        End If
    Next i
End Function

'Function to check if string starts with other string (str starts with start) (no trim)
Public Function StartWith(ByVal str As String, ByVal start As String, Optional ByVal withCase As Boolean = True) As Boolean
    StartWith = IIf(withCase, (Mid(str, 1, Len(start)) = start), (Mid(UCase(str), 1, Len(start)) = UCase(start)))
End Function

'Function to check if string ends with other string (str ends with ending) (no trim)
Public Function EndWith(ByVal str As String, ByVal ending As String, Optional ByVal withCase As Boolean = True) As Boolean
    EndWith = IIf(withCase, (Mid(str, Len(str) - Len(ending) + 1) = ending), (Mid(UCase(str), Len(str) - Len(ending)) = UCase(ending)))
End Function

'Function to add x tabulations at beginning of string
Public Function AddTabs(ByVal str As String, ByVal x As Integer) As String
    While (x > 0)
        str = "    " & str
    Wend
    AddTabs = str
End Function

'Function to get position of first letter in string (returns 0 if there is no letter)
Public Function PosFirstLetter(ByVal str As String) As Long
    Dim chara As Integer
    
    PosFirstLetter = 1
    While (PosFirstLetter <= Len(str))
        chara = Asc(Mid(str, PosFirstLetter, 1))
        If ((chara >= 65 And chara <= 90) Or (chara >= 97 And chara <= 122)) Then Exit Function
        PosFirstLetter = PosFirstLetter + 1
    Wend
    PosFirstLetter = 0
End Function

'Function to get number of occurences of string findStr in string str
Public Function CountOccurences(ByVal str As String, ByVal findStr As String, Optional ByVal withCase As Boolean = True) As Long
    Dim i As Long
    
    If (Not withCase) Then
        str = UCase(str)
        findStr = UCase(findStr)
    End If
    
    For i = 1 To Len(str)
        If (Mid(str, i, Len(findStr)) = findStr) Then CountOccurences = CountOccurences + 1
    Next i
End Function
