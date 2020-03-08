Attribute VB_Name = "StringCheck"
Attribute VB_Description = "Sub and functions to check if a string as a particular format (zip code, phone number, mail address, ...)"
Option Explicit

'Function to check if string is a valid local french phone number (10 consecutive digits, or 10 digits in pairs separated by . 'dots' or spaces)
'Need to activate the reference "Microsoft VBScript Regular Expressions 5.5"
Public Function IsFrenchPhone(ByVal value As String) As Boolean
    Dim reg                             As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^0[1-9](((\.)|( ))(\d\d)){4}$"
    
    IsFrenchPhone = (reg.test(value) Or IsLocalPhone(value))
    Set reg = Nothing
End Function

'Function to check if string is a valid french postal code
'Need to activate the reference "Microsoft VBScript Regular Expressions 5.5"
Public Function IsFrenchPostalCode(ByVal value As String) As Boolean
    Dim reg                             As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^(\d){5}$"
    IsFrenchPostalCode = reg.test(value)
    Set reg = Nothing
End Function

'Function to check if string is a valid international phone number (no specifications for geographics areas and country phone plans)
'Need to activate the reference "Microsoft VBScript Regular Expressions 5.5"
Public Function IsInternationalPhone(ByVal value As String) As Boolean
    Dim reg                             As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^(\+(9[976]\d|8[987530]\d|6[987]\d|5[90]\d|42\d|3[875]\d|2[98654321]\d|" & _
            "9[8543210]|8[6421]|6[6543210]|5[87654321]|4[987654310]|3[9643210]|2[70]|7|1)\d{1,14})$"
    IsInternationalPhone = reg.test(value)
    Set reg = Nothing
End Function

'Function to check if string is a valid local phone number (No formating, only 10 digits)
'Need to activate the reference "Microsoft VBScript Regular Expressions 5.5"
Public Function IsLocalPhone(ByVal value As String) As Boolean
    Dim reg                             As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^0[1-9](\d){8}$"
    IsLocalPhone = reg.test(value)
    Set reg = Nothing
End Function

'Function to check if string is an email address
'Need to activate the reference "Microsoft VBScript Regular Expressions 5.5"
Public Function IsMail(ByVal value As String) As Boolean
    Dim reg                             As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^(\w)+((\w)|(\.\w)|(\-\w)|(_))*@(\w)+((\w)|(\.\w)|(\-\w)|(_))*(\.[\w]{2,3})$"
    IsMail = reg.test(value)
    Set reg = Nothing
End Function

'Function to check if string is a valid phone number
'Using isLocalPhone and isInternationalPhone functions (implemented before)
Public Function IsPhone(ByVal value As String) As Boolean
    IsPhone = (IsInternationalPhone(value) Or IsLocalPhone(value))
End Function

'Function to check if string is a valid local US phone number
'Need to activate the reference "Microsoft VBScript Regular Expressions 5.5"
Public Function IsUSAPhone(ByVal value As String) As Boolean
    Dim reg                             As New VBScript_RegExp_55.RegExp
    
    reg.Pattern = "^([0-9]( |-)?)?(\(?[0-9]{3}\)?|[0-9]{3})( |-)?([0-9]{3}( |-)?[0-9]{4}|[a-zA-Z0-9]{7})$"
    IsUSAPhone = reg.test(value)
    Set reg = Nothing
End Function
