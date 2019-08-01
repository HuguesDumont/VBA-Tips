Attribute VB_Name = "StringDate"
Attribute VB_Description = "Functions for String about date\n"
Option Explicit

'Function to check if the string a correct date (format dd/mm/yyyy HH:mm:ss)
'Differencing leap and non leap years with time on a 24H format
'Dates can be written without the time or time without date (respectively dd/mm/yyyy or HH:mm:ss)
'Need to activate the reference "Microsoft VBScrpt Regular Expressions 5.5"
Public Function IsValidDate(ByVal value As String) As Boolean
    Dim reg As New VBScript_RegExp_55.RegExp

        reg.Pattern = "^(?=\d)(?:(?!(?:(?:0?[5-9]|1[0-4])(?:\.|-|\/)10(?:\.|-|\/)(?:1582))|" & _
                "(?:(?:0?[3-9]|1[0-3])(?:\.|-|\/)0?9(?:\.|-|\/)(?:1752)))(31(?!(?:\.|-|\/)(?:0?[2469]|11))" & _
                "|30(?!(?:\.|-|\/)0?2)|(?:29(?:(?!(?:\.|-|\/)0?2(?:\.|-|\/))|(?=\D0?2\D(?:(?!000[04]|" & _
                "(?:(?:1[^0-6]|[2468][^048]|[3579][^26])00))(?:(?:(?:\d\d)(?:[02468][048]|[13579][26])" & _
                "(?!\x20BC))|(?:00(?:42|3[0369]|2[147]|1[258]|09)\x20BC))))))|2[0-8]|1\d|0?[1-9])([-.\/])" & _
                "(1[012]|(?:0?[1-9]))\2((?=(?:00(?:4[0-5]|[0-3]?\d)\x20BC)|(?:\d{4}(?:$|(?=\x20\d)\x20)))" & _
                "\d{4}(?:\x20BC)?)(?:$|(?=\x20\d)\x20))?((?:(?:0?[1-9]|1[012])(?::[0-5]\d){0,2}(?:\x20[aApP][mM]))" & _
                "|(?:[01]\d|2[0-3])(?::[0-5]\d){1,2})?$"

        IsValidDate = reg.test(value)
        Set reg = Nothing
End Function


