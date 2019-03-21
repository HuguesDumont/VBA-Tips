Attribute VB_Name = "CopyPasteFile"
'Need "Microsoft Scripting Runtime" reference
'Need "OLE Automation" reference
Option Explicit

'Sub to copy and paste file to another directory
Public Sub copierFichier(sourceFolder As String, sourceFile As String, destFolder As String)

    Dim FSO As FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'Verify source file existency
    If Not FSO.FileExists(sourceFolder & sourceFile) Then
        MsgBox "Unable to find source file", vbOKOnly + vbExclamation, "Source file ot found"
    'Verify if there is no file with same name (extension included) in dest folder
    ElseIf Not FSO.FileExists(DossierDest & FichierSrc) Then
        CreateMissingFolder (destFolder)
        FSO.CopyFile (sourceFolder & sourceFile), DossierDest, True
    Else
        MsgBox "File already exists in destation folder", vbExclamation + vbOKOnly, "existing file"
    End If

End Sub

'Create missing folder
Public Sub CreateMissingFolder(folderPath As String)
    If Dir(dossier, vbDirectory) = "" Then MkDir Path:=folderPath
End Sub
