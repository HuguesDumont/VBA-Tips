Attribute VB_Name = "internetFunctions"
Option Explicit

'Sub to navigate to web page in IE (given an url and if the page has to be shown or not)
Public Sub navigateIE(URL As String, view As Boolean)
    Dim IE As Object
    Dim objElement As Object
    Dim objCollection As Object
 
    'Create InternetExplorer Object
    Set IE = CreateObject("InternetExplorer.Application")
 
    'Set IE.Visible = True to make IE visible, or False for IE to run in the background
    IE.Visible = view
    
    'Navigate to URL
    IE.Navigate URL
 
    ' Wait while IE loads
Do While IE.readyState <> 4 And IE.Busy
   DoEvents
Loop
    
    'Clear IE objects
    Set IE = Nothing
    Set objElement = Nothing
    Set objCollection = Nothing
End Sub
