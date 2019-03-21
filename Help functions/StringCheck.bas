Attribute VB_Name = "StringCheck"
Option Explicit

'Function to check if string is an email address
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function isMail(value As String) As Boolean
    Dim reg As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^(\w)+((\w)|(\.\w)|(\-\w)|(_))*@(\w)+((\w)|(\.\w)|(\-\w)|(_))*(\.[\w]{2,3})$"
    isMail = reg.test(value)
    Set reg = Nothing
End Function

'Function to check if string is a valid local phone number (No formating, only 10 digits)
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function isLocalPhone(value As String) As Boolean
    Dim reg As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^0[1-9](\d){8}$"
    isLocalPhone = reg.test(value)
    Set reg = Nothing
End Function

'Function to check if string is a valid local french phone number (10 consecutive digits, or 10 digits in pairs separated by . 'dots' or spaces)
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function isFrenchPhone(value As String) As Boolean
    Dim reg As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^0[1-9](((\.)|( ))(\d\d)){4}$"
    
    isFrenchPhone = (reg.test(value) Or isLocalPhone(value))
    Set reg = Nothing
End Function

'Function to check if string is a valid local US phone number
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function isUSAPhone(value As String) As Boolean
    Dim reg As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^([0-9]( |-)?)?(\(?[0-9]{3}\)?|[0-9]{3})( |-)?([0-9]{3}( |-)?[0-9]{4}|[a-zA-Z0-9]{7})$"
    isUSAPhone = reg.test(value)
    Set reg = Nothing
End Function

'Function to check if string is a valid international phone number (no specifications for geographics areas and country phone plans)
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function isInternationalPhone(value As String) As Boolean
    Dim reg As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^(\+(9[976]\d|8[987530]\d|6[987]\d|5[90]\d|42\d|3[875]\d|2[98654321]\d|" & _
        "9[8543210]|8[6421]|6[6543210]|5[87654321]|4[987654310]|3[9643210]|2[70]|7|1)\d{1,14})$"
    isinternationatlphone = reg.test(value)
    Set reg = Nothing
End Function

'Function to check if string is a valid phone number
'Using isLocalPhone and isInternationalPhone functions (implemented before)
Public Function isPhone(value As String) As Boolean
    isPhone = (isinternationphone(value) Or isLocalPhone(value))
End Function

'Function to check if string is a valid french postal code
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function isFrenchPostalCode(value As String) As Boolean
    Dim reg As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^(\d){5}$"
    isFrenchPostalCode = reg.test(value)
    Set reg = Nothing
End Function
