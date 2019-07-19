Attribute VB_Name = "MathFunctions"
Attribute VB_Description = "Divers math functions such as fibonacci, factorial, GCD, LCM, isPrime, factors, primeFactors, isPerfectNumber, modInverse"
Option Explicit

'Recursive Fibonacci using long type (to avoid memory out of bound)
'Not advised for high numbers (above 25)
'Limit is n = 46
Public Function recursiveFibonacci(ByVal n As Long) As Long
Attribute recursiveFibonacci.VB_Description = "Recursive Fibonacci using long type (to avoid memory out of bound)\r\nNot advised for high numbers (above 25)\r\nLimit is n = 46"
    If n <= 0 Then
        recursiveFibonacci = 0
    ElseIf n = 1 Then
        recursiveFibonacci = 1
    Else
        recursiveFibonacci = recursiveFibonacci(n - 1) + recursiveFibonacci(n - 2)
    End If
End Function

'Iterative Fibonacci using long type (to avoid memory out of bound)
'Limit is n = 46
Function iterativeFibonacci(ByVal n As Long) As Long
Attribute iterativeFibonacci.VB_Description = "Iterative Fibonacci using long type (to avoid memory out of bound)\r\nLimit is n = 46"
    Dim f1 As Long
    Dim f2 As Long
    Dim i As Byte
    
    Select Case n
        Case 0
            iterativeFibonacci = 0
        Case 1, 2
            iterativeFibonacci = 1
        Case Else
            f1 = 1
            f2 = 1
            For i = 3 To n
                iterativeFibonacci = f2 + f1
                f2 = f1
                f1 = iterativeFibonacci
            Next i
   End Select
End Function

'Iterative factorial using long type (to avoid memory out of bound)
'Limit is n = 12
Public Function iterativeFactorial(ByVal n As Long) As Long
Attribute iterativeFactorial.VB_Description = "Iterative factorial using long type (to avoid memory out of bound)\r\nLimit is n = 12"
    Dim cpt As Integer
    
    iterativeFactorial = 1
    For cpt = 1 To n
        iterativeFactorial = cpt * iterativeFactorial
    Next cpt
End Function

'Recursive factorial using long type (to avoid memory out of bound)
'Limit is n = 12
Public Function recursiveFactorial(ByVal n As Long) As Long
Attribute recursiveFactorial.VB_Description = "Recursive factorial using long type (to avoid memory out of bound)\r\nLimit is n = 12"
   If n <= 1 Then
      recursiveFactorial = 1
   Else
      recursiveFactorial = recursiveFactorial(n - 1) * n
   End If
End Function

'GCD (Greatest Common Divisor) for two values using long type
'If a or b <= 0 then return -1
Public Function GCD(ByVal a As Long, ByVal b As Long) As Long
Attribute GCD.VB_Description = "GCD (Greatest Common Divisor) for two values using long type\r\nIf a or b <= 0 then return -1"
    Dim rest As Long
    
    If a <= 0 Or b <= 0 Then
        GCD = -1
        Exit Function
    End If
    
    If a < b Then
        rest = a
        a = b
        b = rest
    End If
    
    rest = a Mod b
    If rest = 0 Then
        GCD = b
    Else
        GCD = GCD(b, rest)
    End If
End Function

'LCM (Least Common Multiple) for two values using long type
'Working with GCD function
'Return -1 if a or b is negative (< 0)
Public Function LCM(ByVal a As Long, ByVal b As Long) As Long
Attribute LCM.VB_Description = "LCM (Least Common Multiple) for two values using long type\r\nWorking with GCD function\r\nReturn -1 if a or b is negative (< 0)"
    If a < 0 Or b < 0 Then
        LCM = -1
    Else
        LCM = a * b / GCD(a, b)
    End If
End Function

'Check if n is a prime number
Public Function isPrime(ByVal n As Long) As Boolean
Attribute isPrime.VB_Description = "Check if n is a prime number"
    Dim i As Double
    If n < 2 Then
        Exit Function
    ElseIf n = 2 Then
        isPrime = True
        Exit Function
    ElseIf Int(n / 2) = (n / 2) Then
        Exit Function
    Else
        For i = 3 To Sqr(n) Step 2
            If Int(n / i) = (n / i) Then
                Exit Function
            End If
        Next i
    End If
    isPrime = True
End Function

'Get previous prime. If n <= 2 then it hasn't a previous prime so function returns -1
Public Function previousPrime(ByVal n As Long) As Long
Attribute previousPrime.VB_Description = "Get previous prime. If n <= 2 then it hasn't a previous prime so function returns -1"
    If (n <= 2) Then
        previousPrime = -1
    ElseIf (n = 3) Then
        previousPrime = 2
    Else
        If (n Mod 2 = 0) Then
            n = n - 1
        Else
            n = n - 2
        End If
        
        While (Not isPrime(n))
            n = n - 2
        Wend
        previousPrime = n
    End If
End Function

'Get next prime number
Public Function nextPrime(ByVal n As Long) As Long
Attribute nextPrime.VB_Description = "Get next prime number"
    If (n < 2) Then
        nextPrime = 2
    Else
        If (n Mod 2 = 0) Then
            n = n + 1
        Else
            n = n + 2
        End If
        
        While (Not isPrime(n))
            n = n + 2
        Wend
        nextPrime = n
    End If
End Function

'Function to get all factors of a number (returns a string)
Public Function factors(ByVal n As Long) As String
Attribute factors.VB_Description = "Function to get all factors of a number (returns a string)"
    Dim i As Long
    Dim corresponding As String
    
    If n < 1 Then
        MsgBox "Value cannot be below 1.", vbOKOnly + vbExclamation, "Invalid number"
        Exit Function
    End If
    
    factors = 1
    corresponding = n
    
    For i = 2 To Sqr(n)
        If n Mod i = 0 Then
            factors = factors & ", " & i
            If i <> n / i Then
                corresponding = n / i & ", " & corresponding
            End If
        End If
    Next i
    
    If n <> 1 Then
        factors = factors & ", " & corresponding
    End If
End Function

'Function to get prime factors of a number (returns a string)
'Works with isPrime function
Public Function primeFactors(ByVal n As Long) As String
Attribute primeFactors.VB_Description = "Function to get prime factors of a number (returns a string)\r\nWorks with isPrime function"
    Dim i As Long
    Dim corresponding As String
    
    If n < 2 Then
        MsgBox "Value cannot be below 0.", vbOKOnly + vbExclamation, "Invalid number"
        Exit Function
    End If
    
    If isPrime(n) Then
        primeFactors = n
        Exit Function
    End If
    
    If n Mod 2 = 0 Then
        primeFactors = "2"
        If (isPrime(n / 2) And n <> 4) Then
            corresponding = CStr(n / 2)
        End If
    End If
    
    For i = 3 To Int(n / 2) + 1 Step 2
        If n Mod i = 0 And isPrime(i) Then
            primeFactors = IIf(primeFactors <> "", primeFactors & ", " & CStr(i), CStr(i))
            If ((i <> (n / i)) And isPrime(n / i)) Then
                corresponding = IIf(corresponding <> "", CStr(n / i) & ", " & corresponding, CStr(n / i))
            End If
        End If
    Next i
    
    If corresponding <> "" Then
        primeFactors = primeFactors & ", " & corresponding
    End If
End Function

'Function to check if a number is a perfect number
'Works with factors, isPrime and sumLongArray functions
Public Function isPerfectNumber(ByVal n As Long) As Boolean
Attribute isPerfectNumber.VB_Description = "Function to check if a number is a perfect number\r\nWorks with factors, isPrime and sumLongArray functions"
    Dim factorsArray() As Long
    Dim i As Long
    Dim splitting() As String
    
    If (n < 4 Or (n Mod 2 = 1) Or isPrime(n)) Then
        isPerfectNumber = False
        Exit Function
    End If
    
    splitting = Split(factors(n), ", ")
    
    ReDim factorsArray(UBound(splitting))
    
    For i = 0 To UBound(splitting) - 1
        factorsArray(i) = val(splitting(i))
    Next i
    isPerfectNumber = ((sumLongArray(factorsArray)) = n)
End Function

'Function to sum all values in long array
Public Function sumLongArray(arr() As Long) As Long
Attribute sumLongArray.VB_Description = "Function to sum all values in long array"
    Dim i As Long
    
    sumLongArray = 0
    For i = 0 To UBound(arr)
        sumLongArray = sumLongArray + arr(i)
    Next i
End Function

'Function to calculate the modular inverse (x mod inverse n)
'If x isn't invertible then return -1
Public Function modInverse(ByVal x As Long, ByVal n As Long) As Long
Attribute modInverse.VB_Description = "Function to calculate the modular inverse (x mod inverse n)\r\nIf x isn't invertible then return -1"
    Dim t As Long, nt As Long, r As Long, nr As Long, q As Long, tmp As Long

    If n < 0 Then
        n = -n
    End If
    
    If x < 0 Then
        x = n - ((-x) Mod n)
    End If
    
    t = 0
    nt = 1
    r = n
    nr = x
    
    While nr <> 0
        q = r / nr
        tmp = t
        t = nt
        nt = tmp - q * nt
        tmp = r
        r = nr
        nr = tmp - q * nr
    Wend
    
    If r > 1 Then
        modInverse = -1
    Else
        If t < 0 Then
            t = t + n
        End If
        modInverse = t
    End If
End Function

'Function to check if number a is a divisor of number b
Public Function isDivisor(ByVal a As Long, ByVal b As Long) As Boolean
Attribute isDivisor.VB_Description = "Check if first number is a divisor of second number (using long types)"
    isDivisor = IIf(b Mod a = 0, True, False)
End Function

'Function to sum all digits of a number together once
Public Function sumDigitsOnce(ByVal n As Long) As Long
Attribute sumDigitsOnce.VB_Description = "Function to sum all digits of a number together once"
    Dim str As String
    Dim sum As Long
    Dim i As Long
    
    str = CStr(n)
    sum = 0
    For i = 1 To Len(str)
        sum = sum + CLng(Mid(str, i, 1))
    Next i
    sumDigitsOnce = sum
End Function

'Function to sum all digits of a number together until one digit is left
Public Function sumAllDigits(ByVal n As Long) As Integer
Attribute sumAllDigits.VB_Description = "Function to sum all digits of a number together until one digit is left"
    Dim str As String
    Dim sum As Long
    Dim i As Long
    
    str = CStr(n)
    If (Len(str) = 1) Then
        sumAllDigits = n
    Else
        sum = 0
        For i = 1 To Len(str)
            sum = sum + CLng(Mid(str, i, 1))
        Next i
        sumAllDigits = sumAllDigits(sum)
    End If
End Function

'Factorial inverse of n. If n isn't a factorial then return -1
Public Function factorialInverse(ByVal n As Long) As Long
Attribute factorialInverse.VB_Description = "Factorial inverse of n. If n isn't a factorial then return -1"
    Dim i As Long
    Dim res As Double
    
    i = 1
    res = CDbl(n)
    While (res > 1 And res = Int(res))
        i = i + 1
        res = res / i
    Wend
    factorialInverse = IIf(res = Int(res), i, -1)
End Function
