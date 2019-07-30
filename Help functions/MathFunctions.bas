Attribute VB_Name = "MathFunctions"
Attribute VB_Description = "Divers math functions such as fibonacci, factorial, GCD, LCM, isPrime, factors, primeFactors, isPerfectNumber, modInverse"
Option Explicit

'Function to get all common divisors of a and b. If a or b < 1 returns empty array
Public Function CommonDivisors(ByVal a As Long, ByVal b As Long) As Variant
Attribute CommonDivisors.VB_Description = "Function to get all common divisors of a and b. If a or b < 1 returns empty array"
    Dim i As Long, tmp As Long, count As Long
    Dim tmpArray As Variant
    
    If (a > b) Then
        tmp = a
        a = b
        b = tmp
    End If
    
    tmpArray = Array()
    If (a >= 1 And b >= 1) Then
        ReDim tmpArray(a)
        tmpArray(0) = 1
        count = 0
        For i = 2 To a
            If (IsDivisor(i, a) And IsDivisor(i, b)) Then
                count = count + 1
                tmpArray(count) = i
            End If
        Next i
        ReDim Preserve tmpArray(count)
    End If
    CommonDivisors = tmpArray
End Function

'Function to get all divisors of n (if n < 1 then return empty array)
Public Function Divisors(ByVal n As Long) As Variant
Attribute Divisors.VB_Description = "Function to get all divisors of n (if n < 1 then return empty array)"
    Dim i As Long, count As Long
    Dim tmpArray As Variant
    
    tmpArray = Array()
    If (n >= 1) Then
        ReDim tmpArray(n)
        tmpArray(0) = 1
        count = 1
        For i = 2 To n
            If (n Mod i = 0) Then
                tmpArray(count) = i
                count = count + 1
            End If
        Next i
        ReDim Preserve tmpArray(count - 1)
    End If
    Divisors = tmpArray
End Function

'Recursive Fibonacci using long type (to avoid memory out of bound)
'Not advised for high numbers (above 25)
'Limit is n = 46
Public Function RecursiveFibonacci(ByVal n As Long) As Long
Attribute RecursiveFibonacci.VB_Description = "Recursive Fibonacci using long type (to avoid memory out of bound)\r\nNot advised for high numbers (above 25)\r\nLimit is n = 46"
    If n <= 0 Then
        RecursiveFibonacci = 0
    ElseIf n = 1 Then
        RecursiveFibonacci = 1
    Else
        RecursiveFibonacci = RecursiveFibonacci(n - 1) + RecursiveFibonacci(n - 2)
    End If
End Function

'Iterative Fibonacci using long type (to avoid memory out of bound)
'Limit is n = 46
Function IterativeFibonacci(ByVal n As Long) As Long
Attribute IterativeFibonacci.VB_Description = "Iterative Fibonacci using long type (to avoid memory out of bound)\r\nLimit is n = 46"
    Dim f1 As Long
    Dim f2 As Long
    Dim i As Byte
    
    Select Case n
        Case 0
            IterativeFibonacci = 0
        Case 1, 2
            IterativeFibonacci = 1
        Case Else
            f1 = 1
            f2 = 1
            For i = 3 To n
                IterativeFibonacci = f2 + f1
                f2 = f1
                f1 = IterativeFibonacci
            Next i
   End Select
End Function

'Iterative factorial using long type (to avoid memory out of bound)
'Limit is n = 12
Public Function IterativeFactorial(ByVal n As Long) As Long
Attribute IterativeFactorial.VB_Description = "Iterative factorial using long type (to avoid memory out of bound)\r\nLimit is n = 12"
    Dim cpt As Integer
    
    IterativeFactorial = 1
    For cpt = 1 To n
        IterativeFactorial = cpt * IterativeFactorial
    Next cpt
End Function

'Recursive factorial using long type (to avoid memory out of bound)
'Limit is n = 12
Public Function RecursiveFactorial(ByVal n As Long) As Long
Attribute RecursiveFactorial.VB_Description = "Recursive factorial using long type (to avoid memory out of bound)\r\nLimit is n = 12"
   If n <= 1 Then
      RecursiveFactorial = 1
   Else
      RecursiveFactorial = RecursiveFactorial(n - 1) * n
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
Public Function IsPrime(ByVal n As Long) As Boolean
Attribute IsPrime.VB_Description = "Check if n is a prime number"
    Dim i As Double
    If n < 2 Then
        Exit Function
    ElseIf n = 2 Then
        IsPrime = True
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
    IsPrime = True
End Function

'Get previous prime. If n <= 2 then it hasn't a previous prime so function returns -1
Public Function PreviousPrime(ByVal n As Long) As Long
Attribute PreviousPrime.VB_Description = "Get previous prime. If n <= 2 then it hasn't a previous prime so function returns -1"
    If (n <= 2) Then
        PreviousPrime = -1
    ElseIf (n = 3) Then
        PreviousPrime = 2
    Else
        If (n Mod 2 = 0) Then
            n = n - 1
        Else
            n = n - 2
        End If
        
        While (Not IsPrime(n))
            n = n - 2
        Wend
        PreviousPrime = n
    End If
End Function

'Get next prime number
Public Function NextPrime(ByVal n As Long) As Long
Attribute NextPrime.VB_Description = "Get next prime number"
    If (n < 2) Then
        NextPrime = 2
    Else
        If (n Mod 2 = 0) Then
            n = n + 1
        Else
            n = n + 2
        End If
        
        While (Not IsPrime(n))
            n = n + 2
        Wend
        NextPrime = n
    End If
End Function

'Returns an array containing the factors of n. If n < 1 then returns empty array.
Public Function Factors(ByVal n As Long) As Variant
Attribute Factors.VB_Description = "Returns an array containing the factors of n. If n < 1 then returns empty array."
    Dim i As Long, count As Long, count2 As Long
    Dim tmpArray As Variant, corresponding As Variant
    
    tmpArray = Array()
    If n >= 1 Then
        count = 0
        ReDim tmpArray(Int(n / 2) + 1)
        tmpArray(0) = 1
        
        If (n > 1) Then
            count = 1
            count2 = 1
            ReDim corresponding(Int(n / 2) + 1)
            corresponding(0) = n
            
            For i = 2 To Int(Sqr(n))
                If n Mod i = 0 Then
                    tmpArray(count) = i
                    count = count + 1
                    If i <> n / i Then
                        corresponding(count2) = n / i
                        count2 = count2 + 1
                    End If
                End If
            Next i
            
             If (Not IsEmpty(corresponding(0))) Then
                For i = count2 - 1 To 0
                    tmpArray(count) = corresponding(i)
                    count = count + 1
                Next i
            End If
            ReDim Preserve tmpArray(count - 1)
        End If
    End If
    Factors = tmpArray
End Function

'Function to get prime factors of a number (if n < 2 then return empty array)
'Works with isPrime function
Public Function PrimeFactors(ByVal n As Long) As Variant
Attribute PrimeFactors.VB_Description = "Function to get prime factors of a number (if n < 2 then return empty array)\r\nWorks with isPrime function"
    Dim i As Long, count As Long, count2 As Long
    Dim tmpArray As Variant, corresponding As Variant
    
    tmpArray = Array()
    If (n >= 2) Then
        If IsPrime(n) Then
            ReDim tmpArray(1)
            tmpArray(0) = n
        Else
            ReDim corresponding(Int(n / 2) + 1)
            ReDim tmpArray(Int(n / 2) + 1)
            count = 0
            count2 = 0
            If n Mod 2 = 0 Then
                tmpArray(0) = 2
                count = 1
                If (IsPrime(n / 2) And n <> 4) Then
                    corresponding(0) = n / 2
                    count2 = 1
                End If
            End If
            
            For i = 3 To Int(n / 2) + 1 Step 2
                If n Mod i = 0 And IsPrime(i) Then
                    tmpArray(count) = i
                    If ((i <> (n / i)) And IsPrime(n / i)) Then
                        corresponding(count2) = n / i
                        count2 = count2 + 1
                    End If
                    count = count + 1
                End If
            Next i
            
            If (Not IsEmpty(corresponding(0))) Then
                For i = count2 - 1 To 0
                    tmpArray(count) = corresponding(i)
                    count = count + 1
                Next i
            End If
            ReDim Preserve tmpArray(count - 1)
        End If
    End If
    PrimeFactors = tmpArray
End Function

'Function to check if a number is a perfect number
'Works with factors, isPrime and sumLongArray functions
Public Function IsPerfectNumber(ByVal n As Long) As Boolean
Attribute IsPerfectNumber.VB_Description = "Function to check if a number is a perfect number\r\nWorks with factors, isPrime and sumLongArray functions"
    Dim factorsArray() As Long
    Dim i As Long
    Dim splitting() As String
    
    If (n < 4 Or (n Mod 2 = 1) Or IsPrime(n)) Then
        IsPerfectNumber = False
        Exit Function
    End If
    
    splitting = Split(Factors(n), ", ")
    
    ReDim factorsArray(UBound(splitting))
    
    For i = 0 To UBound(splitting) - 1
        factorsArray(i) = val(splitting(i))
    Next i
    IsPerfectNumber = ((SumLongArray(factorsArray)) = n)
End Function

'Function to sum all values in long array
Public Function SumLongArray(arr() As Long) As Long
Attribute SumLongArray.VB_Description = "Function to sum all values in long array"
    Dim i As Long
    
    SumLongArray = 0
    For i = 0 To UBound(arr)
        SumLongArray = SumLongArray + arr(i)
    Next i
End Function

'Function to calculate the modular inverse (x mod inverse n)
'If x isn't invertible then return -1
Public Function ModInverse(ByVal x As Long, ByVal n As Long) As Long
Attribute ModInverse.VB_Description = "Function to calculate the modular inverse (x mod inverse n)\r\nIf x isn't invertible then return -1"
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
        ModInverse = -1
    Else
        If t < 0 Then
            t = t + n
        End If
        ModInverse = t
    End If
End Function

'Function to check if number a is a divisor of number b
Public Function IsDivisor(ByVal a As Long, ByVal b As Long) As Boolean
Attribute IsDivisor.VB_Description = "Check if first number is a divisor of second number (using long types)"
    IsDivisor = IIf(b Mod a = 0, True, False)
End Function

'Function to sum all digits of a number together once
Public Function SumDigitsOnce(ByVal n As Long) As Long
Attribute SumDigitsOnce.VB_Description = "Function to sum all digits of a number together once"
    Dim str As String
    Dim sum As Long
    Dim i As Long
    
    str = CStr(n)
    sum = 0
    For i = 1 To Len(str)
        sum = sum + CLng(Mid(str, i, 1))
    Next i
    SumDigitsOnce = sum
End Function

'Function to sum all digits of a number together until one digit is left
Public Function SumAllDigits(ByVal n As Long) As Integer
Attribute SumAllDigits.VB_Description = "Function to sum all digits of a number together until one digit is left"
    Dim str As String
    Dim sum As Long
    Dim i As Long
    
    str = CStr(n)
    If (Len(str) = 1) Then
        SumAllDigits = n
    Else
        sum = 0
        For i = 1 To Len(str)
            sum = sum + CLng(Mid(str, i, 1))
        Next i
        SumAllDigits = SumAllDigits(sum)
    End If
End Function

'Factorial inverse of n. If n isn't a factorial then return -1
Public Function FactorialInverse(ByVal n As Long) As Long
Attribute FactorialInverse.VB_Description = "Factorial inverse of n. If n isn't a factorial then return -1"
    Dim i As Long
    Dim res As Double
    
    i = 1
    res = CDbl(n)
    While (res > 1 And res = Int(res))
        i = i + 1
        res = res / i
    Wend
    FactorialInverse = IIf(res = Int(res), i, -1)
End Function

'Function to get the n first multiples of number (number and n have to be over 1, else returns empty array)
Public Function FirstMultiples(ByVal number As Long, ByVal n As Long) As Variant
Attribute FirstMultiples.VB_Description = "Function to get the n first multiples of number (number and n have to be over 1, else returns empty array)"
    Dim i As Long
    Dim tmpArray As Variant
    
    tmpArray = Array()
    If (number > 1 And n > 1) Then
        ReDim tmpArray(n - 1)
        tmpArray(0) = number
        For i = 2 To n
            tmpArray(i - 1) = number * i
        Next i
    End If
    FirstMultiples = tmpArray
End Function

'Function to get the n first multiples of number (x, y and n have to be over 1, else returns empty array)
'Needs LCM function
Public Function FirstCommonMultiples(ByVal x As Long, ByVal y As Long, ByVal n As Long) As Variant
Attribute FirstCommonMultiples.VB_Description = "Function to get the n first multiples of number (x, y and n have to be over 1, else returns empty array)\r\nNeeds LCM function"
    Dim i As Long
    Dim leastCM As Long
    Dim tmpArray As Variant
    
    tmpArray = Array()
    If (x > 1 And y > 1 And n > 1) Then
        ReDim tmpArray(n - 1)
        leastCM = LCM(x, y)
        tmpArray(0) = leastCM
        For i = 2 To n
            tmpArray(i - 1) = leastCM * i
        Next i
    End If
    FirstCommonMultiples = tmpArray
End Function
