Attribute VB_Name = "MathFunctions"
Option Explicit

'Recursive Fibonacci using long type (to avoid memory out of bound)
'Not advised for high numbers (above 25)
'Limit is n = 46
Public Function recursiveFibonacci(n As Long) As Long
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
Function iterativeFibonacci(n As Byte) As Long
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
Public Function iterativeFactorial(n As Long) As Long
    Dim cpt As Integer
    
    iterativeFactorial = 1
    For cpt = 1 To n
        iterativeFactorial = cpt * iterativeFactorial
    Next cpt
End Function

'Iterative factorial using long type (to avoid memory out of bound)
'Limit is n = 12
Public Function recursiveFactorial(n As Long) As Long
   If n <= 1 Then
      recursiveFactorial = 1
   Else
      recursiveFactorial = recursiveFactorial(n - 1) * n
   End If
End Function

'GCD (Greatest Common Divisor) for two values using long type
'If a or b <= 0 then return -1
Public Function GCD(a As Long, b As Long) As Long
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
Public Function LCM(a As Long, b As Long) As Long
    If a < 0 Or b < 0 Then
        LCM = -1
    Else
        LCM = a * b / GCD(a, b)
    End If
End Function

'Check if n is a prime number
Public Function isPrime(n As Long) As Boolean
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

'Function to get all factors of a number (in a string)
Public Function factors(n As Long) As String
    Dim i As Long
    Dim corresponding As String
    
    If n < 1 Then
        MsgBox "Value cannot be below 0.", vbOKOnly + vbExclamation, "Invalid number"
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

'Function to get prime factors of a number (in a string)
'Works with isPrime function
Public Function primeFactors(n As Long) As String
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
    
    For i = 3 To Sqr(n) Step 2
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
Public Function isPerfectNumber(n As Long) As Boolean
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
        factorsArray(i) = Val(splitting(i))
    Next i
    isPerfectNumber = ((sumLongArray(factorsArray)) = n)
End Function

'Function to sum all values in long array
Public Function sumLongArray(arr() As Long) As Long
    Dim i As Long
    
    sumLongArray = 0
    For i = 0 To UBound(arr) - 1
        sumLongArray = sumLongArray + arr(i)
    Next i
End Function
