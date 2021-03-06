Attribute VB_Name = "Base64"
Attribute VB_Description = "Encode and decode for Base64"
Option Explicit

'I don't own the credits for these two functions. They both come from here : http://www.vbforums.com/showthread.php?t=379072

Private Const CLONEMASK As Long = 16515072  '000000 111111 111111 111111
Private Const CLTWOMASK As Long = 258048    '111111 000000 111111 111111
Private Const CLTHREEMASK As Long = 4032    '111111 111111 000000 111111
Private Const CLFOURMASK As Long = 63       '111111 111111 111111 000000

Private Const CLHIGHMASK As Long = 16711680 '11111111 00000000 00000000
Private Const CLMIDMASK As Long = 65280     '00000000 11111111 00000000
Private Const CLLOWMASK As Long = 255       '00000000 00000000 11111111

Private Const CL2EXP18 As Long = 262144     '2 to the 18th power
Private Const CL2EXP12 As Long = 4096       '2 to the 12th
Private Const CL2EXP6 As Long = 64          '2 to the 6th
Private Const CL2EXP8 As Long = 256         '2 to the 8th
Private Const CL2EXP16 As Long = 65536      '2 to the 16th

'Encode string to Base64
Public Function Encode64(ByVal sString As String) As String
    Dim bTrans(63)                      As Byte
    Dim lPowers8(255)                   As Long
    Dim lPowers16(255)                  As Long
    Dim bOut()                          As Byte
    Dim bIn()                           As Byte
    Dim lChar                           As Long
    Dim lTrip                           As Long
    Dim iPad                            As Integer
    Dim lLen                            As Long
    Dim lTemp                           As Long
    Dim lPos                            As Long
    Dim lOutSize                        As Long
    
    For lTemp = 0 To 63                                 'Fill the translation table.
        Select Case lTemp
            Case 0 To 25
                bTrans(lTemp) = 65 + lTemp              'A - Z
            Case 26 To 51
                bTrans(lTemp) = 71 + lTemp              'a - z
            Case 52 To 61
                bTrans(lTemp) = lTemp - 4               '1 - 0
            Case 62
                bTrans(lTemp) = 43                      'Chr(43) = "+"
            Case 63
                bTrans(lTemp) = 47                      'Chr(47) = "/"
        End Select
    Next lTemp
    
    For lTemp = 0 To 255                                'Fill the 2^8 and 2^16 lookup tables.
        lPowers8(lTemp) = lTemp * CL2EXP8
        lPowers16(lTemp) = lTemp * CL2EXP16
    Next lTemp
    
    iPad = Len(sString) Mod 3                           'See if the length is divisible by 3
    
    If iPad Then                                        'If not, figure out the end pad and resize the input.
        iPad = 3 - iPad
        sString = sString & String(iPad, Chr(0))
    End If
    
    bIn = strConv(sString, vbFromUnicode)               'Load the input string.
    lLen = ((UBound(bIn) + 1) \ 3) * 4                  'Length of resulting string.
    lTemp = lLen \ 72                                   'Added space for vbCrLfs.
    lOutSize = ((lTemp * 2) + lLen) - 1                 'Calculate the size of the output buffer.
    ReDim bOut(lOutSize)                                'Make the output buffer.
    
    lLen = 0                                            'Reusing this one, so reset it.
    
    For lChar = LBound(bIn) To UBound(bIn) Step 3
        lTrip = lPowers16(bIn(lChar)) + lPowers8(bIn(lChar + 1)) + bIn(lChar + 2)    'Combine the 3 bytes
        lTemp = lTrip And CLONEMASK                     'Mask for the first 6 bits
        bOut(lPos) = bTrans(lTemp \ CL2EXP18)           'Shift it down to the low 6 bits and get the value
        lTemp = lTrip And CLTWOMASK                     'Mask for the second set.
        bOut(lPos + 1) = bTrans(lTemp \ CL2EXP12)       'Shift it down and translate.
        lTemp = lTrip And CLTHREEMASK                   'Mask for the third set.
        bOut(lPos + 2) = bTrans(lTemp \ CL2EXP6)        'Shift it down and translate.
        bOut(lPos + 3) = bTrans(lTrip And CLFOURMASK)   'Mask for the low set.
        
        If (lLen = 68) Then                               'Ready for a newline
            bOut(lPos + 4) = 13                         'Chr(13) = vbCr
            bOut(lPos + 5) = 10                         'Chr(10) = vbLf
            lLen = 0                                    'Reset the counter
            lPos = lPos + 6
        Else
            lLen = lLen + 4
            lPos = lPos + 4
        End If
    Next lChar
    
    If (bOut(lOutSize) = 10) Then lOutSize = lOutSize - 2 'Shift the padding chars down if it ends with CrLf.
    
    If (iPad = 1) Then                                    'Add the padding chars if any.
        bOut(lOutSize) = 61                             'Chr(61) = "="
    ElseIf (iPad = 2) Then
        bOut(lOutSize) = 61
        bOut(lOutSize - 1) = 61
    End If
    
    Encode64 = strConv(bOut, vbUnicode)                 'Convert back to a string and return it.
End Function

'Decode Base64 to string
Public Function Decode64(ByVal sString As String) As String
    Dim bOut()                          As Byte
    Dim bIn()                           As Byte
    Dim bTrans(255)                     As Byte
    Dim lPowers6(63)                    As Long
    Dim lPowers12(63)                   As Long
    Dim lPowers18(63)                   As Long
    Dim lQuad                           As Long
    Dim iPad                            As Integer
    Dim lChar                           As Long
    Dim lPos                            As Long
    Dim sOut                            As String
    Dim lTemp                           As Long
    
    sString = Replace(sString, vbCr, vbNullString)      'Get rid of the vbCrLfs.  These could be in...
    sString = Replace(sString, vbLf, vbNullString)      'either order.
    
    lTemp = Len(sString) Mod 4                          'Test for valid input.
    
    If lTemp Then
        Call Err.Raise(vbObjectError, "MyDecode", "Input string is not valid Base64.")
    End If
    
    If InStrRev(sString, "==") Then                     'InStrRev is faster when you know it's at the end.
        iPad = 2                                        'Note:  These translate to 0, so you can leave them...
    ElseIf InStrRev(sString, "=") Then                  'in the string and just resize the output.
        iPad = 1
    End If
    
    For lTemp = 0 To 255                                'Fill the translation table.
        Select Case lTemp
            Case 65 To 90
                bTrans(lTemp) = lTemp - 65              'A - Z
            Case 97 To 122
                bTrans(lTemp) = lTemp - 71              'a - z
            Case 48 To 57
                bTrans(lTemp) = lTemp + 4               '1 - 0
            Case 43
                bTrans(lTemp) = 62                      'Chr(43) = "+"
            Case 47
                bTrans(lTemp) = 63                      'Chr(47) = "/"
        End Select
    Next lTemp
    
    For lTemp = 0 To 63                                 'Fill the 2^6, 2^12, and 2^18 lookup tables.
        lPowers6(lTemp) = lTemp * CL2EXP6
        lPowers12(lTemp) = lTemp * CL2EXP12
        lPowers18(lTemp) = lTemp * CL2EXP18
    Next lTemp
    
    bIn = strConv(sString, vbFromUnicode)               'Load the input byte array.
    ReDim bOut((((UBound(bIn) + 1) \ 4) * 3) - 1)       'Prepare the output buffer.
    
    For lChar = 0 To UBound(bIn) Step 4
        lQuad = lPowers18(bTrans(bIn(lChar))) + lPowers12(bTrans(bIn(lChar + 1))) + _
                lPowers6(bTrans(bIn(lChar + 2))) + bTrans(bIn(lChar + 3))           'Rebuild the bits.
        lTemp = lQuad And CLHIGHMASK                    'Mask for the first byte
        bOut(lPos) = lTemp \ CL2EXP16                   'Shift it down
        lTemp = lQuad And CLMIDMASK                     'Mask for the second byte
        bOut(lPos + 1) = lTemp \ CL2EXP8                'Shift it down
        bOut(lPos + 2) = lQuad And CLLOWMASK            'Mask for the third byte
        lPos = lPos + 3
    Next lChar
    
    sOut = strConv(bOut, vbUnicode)                     'Convert back to a string.
    
    If iPad Then sOut = Left$(sOut, Len(sOut) - iPad)   'Chop off any extra bytes.
    Decode64 = sOut
End Function
