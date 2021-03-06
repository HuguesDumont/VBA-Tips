Attribute VB_Name = "FileAndFolderFunctions"
Attribute VB_Description = "Functions to handle file and folder processing\r\nNeed ""Microsoft Scripting Runtime"" reference\r\nNeed ""OLE Automation"" reference"
Option Explicit

'Need "Microsoft Scripting Runtime" reference
'Need "OLE Automation" reference

'Sub to copy all files from one directory to another
'If dest folder doesn't exists then creates it
'If destName <> "" then copy will have a different name from source
Public Sub CopyAllFilesTo(ByVal fName As String, ByVal destpath As String, Optional ByVal sourcePath As String = vbNullString, Optional ByVal destName As String = vbNullString, Optional ByVal overwrite As Boolean = False)
    Dim FSO                             As New Scripting.FileSystemObject
    
    fName = FormatPath(fName, sourcePath)
    
    If FolderExists(fName) Then
        destpath = FormatPath(destName, destpath)
        CreateFolder destpath
        Call FSO.copyFile(fName & "\*.*", destpath, (overwrite Or Not FolderExists(destpath)))
    Else
        MsgBox "Source folder not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot copy folder"
    End If
End Sub

'Sub to copy a file to another folder
'If dest folder doesn't exists then creates it
'If destName <> "" then copy will have a different name from source
Public Sub CopyFileTo(ByVal fName As String, ByVal destpath As String, Optional ByVal sourcePath As String = vbNullString, Optional ByVal destName As String = vbNullString, Optional ByVal overwrite As Boolean = False)
    fName = FormatPath(fName, sourcePath)
    
    If FileExists(fName) Then
        destpath = IIf(destName <> vbNullString, FormatPath(destName, destpath), FormatPath(GetSourceName(fName), destpath))
        
        If (overwrite Or (Not FileExists(destpath))) Then
            CreateFolder GetFolderName(destpath)
            FileCopy fName, destpath
        Else
            MsgBox "File already exists in destination folder.", vbOKOnly + vbExclamation, "File already exists"
        End If
    Else
        MsgBox "Source file not found. Please check your source file path.", vbOKOnly + vbExclamation, "Cannot copy file"
    End If
End Sub

'Sub to copy a folder to another folder
'If dest folder doesn't exists then creates it
'If destName <> "" then copy will have a different name from source
Public Sub CopyFolderTo(ByVal fName As String, ByVal destpath As String, Optional ByVal sourcePath As String = vbNullString, Optional ByVal destName As String = vbNullString, Optional ByVal overwrite As Boolean = False)
    Dim FSO                             As New Scripting.FileSystemObject
    
    fName = FormatPath(fName, sourcePath)
    
    If FolderExists(fName) Then
        destpath = IIf(destName <> vbNullString, FormatPath(destName, destpath), FormatPath(GetSourceName(fName), destpath))
        Call FSO.CopyFolder(fName, destpath, (overwrite Or Not FolderExists(destpath)))
    Else
        MsgBox "Source folder not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot copy folder"
    End If
End Sub

'Sub to copy file or folder to another folder
'If dest folder doesn't exists then creates it
'If destName <> "" then copy will have a different name from source
Public Sub CopyTo(ByVal fName As String, ByVal destpath As String, Optional ByVal sourcePath As String = vbNullString, Optional ByVal destName As String = vbNullString, Optional ByVal overwrite As Boolean = False)
    If FolderExists(FormatPath(fName, sourcePath)) Then
        Call CopyFolderTo(fName, destpath, sourcePath, destName, overwrite)
    ElseIf FileExists(FormatPath(fName, sourcePath)) Then
        Call CopyFileTo(fName, destpath, sourcePath, destName, overwrite)
    Else
        MsgBox "Source not found. Please check your source path.", vbOKOnly + vbExclamation, "Cannot copy"
    End If
End Sub

'Sub to create file (using name and path separately)
'If dest = "" then fName is full path
'Do not overwrite by default
Public Sub CreateFile(ByVal fName As String, Optional ByVal dest As String = vbNullString, Optional ByVal overwrite As Boolean = False)
    Dim FSO                             As New Scripting.FileSystemObject
    Dim oFile                           As Object
    
    On Error GoTo cantCreate
    
    'Get full path
    fName = FormatPath(fName, dest)
    
    'Create folder if necessary
    CreateFolder GetFolderName(fName)
    
    'Create the file
    Set oFile = FSO.CreateTextFile(fName, overwrite)
    oFile.Close
    
    Set oFile = Nothing
    Set FSO = Nothing
    Exit Sub
    
cantCreate:
    MsgBox "Error while trying to create new file. Please check your path and file name don't contain incompatible characters or if you have the rights to create the file.", _
            vbOKOnly + vbCritical, "Error creating file"
    Set oFile = Nothing
    Set FSO = Nothing
End Sub

'Sub to create folder (if folder doesn't exist)
'If dest = "" then fName is full path
Public Sub CreateFolder(ByVal fName As String, Optional ByVal dest As String = vbNullString)
    On Error GoTo cantCreate
    
    fName = FormatPath(fName, dest)
    
    If (Not FolderExists(fName)) Then
        Call CreateFolder(GetFolderName(fName))
        MkDir fName
    End If
    
    Exit Sub
    
cantCreate:
    MsgBox "Error while trying to create new folder. Please check your path and folder name don't contain incompatible characters or if you have the rights to create the folder.", _
            vbOKOnly + vbCritical, "Error creating folder"
End Sub

'Sub to delete all files in folder
'If fPath = "" then fName is full path
Public Sub DeleteAllFiles(ByVal fName As String, Optional ByVal fPath As String = vbNullString)
    On Error GoTo errorDeleting
    
    If FolderExists(FormatPath(fName, fPath)) Then
        Kill FormatPath(fName, fPath) & "*.*"
    Else
        MsgBox "Source folder not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot delete files"
    End If
    
errorDeleting:
    MsgBox "Error while deleting files. Please check that your files are closed and you have the correct rigths.", vbOKOnly + vbCritical, "Cannot delete files"
End Sub

'Sub to delete file
'If fPath = "" then fName is full Path
Public Sub DeleteFile(ByVal fName As String, Optional ByVal fPath As String = vbNullString)
    On Error GoTo errorDeleting
    
    If FileExists(FormatPath(fName, fPath)) Then
        Kill FormatPath(fName, fPath)
    Else
        MsgBox "Source file not found. Please check your source file path.", vbOKOnly + vbExclamation, "Cannot delete file"
    End If
    
errorDeleting:
    MsgBox "Error while deleting file. Please check that your file is closed and you have the correct rigths.", vbOKOnly + vbCritical, "Cannot delete file"
End Sub

'Sub to delete folder
'If fPath = "" then fName is full Path
Public Sub DeleteFolder(ByVal fName As String, Optional ByVal fPath As String = vbNullString)
    Dim FSO                             As New Scripting.FileSystemObject
    
    fName = FormatPath(fName, fPath)
    
    If FolderExists(fName) Then
        On Error Resume Next
        
        'Delete files
        FSO.DeleteFile fName & "\*.*", True
        
        'Delete subfolders
        FSO.DeleteFolder fName & "\*.*", True
        
        Set FSO = Nothing
        
        RmDir fName
        
        On Error GoTo 0
    Else
        MsgBox "Source folder not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot delete folder"
    End If
End Sub

'Sub to delete folder or file
'If fPath = "" then fName is full Path
Public Sub Deletion(ByVal fName As String, Optional ByVal fPath As String = vbNullString)
    On Error Resume Next
    DeleteFolder fName, fPath
    DeleteFile fName, fPath
    On Error GoTo 0
End Sub

'Function to check if file or folder exists
Public Function FExists(ByVal fName As String, Optional ByVal fPath As String = vbNullString) As Boolean
    FExists = (Len(Dir(FormatPath(fName, fPath), vbDirectory)) <> 0)
End Function

'Function to check if file exists
Public Function FileExists(ByVal fName As String, Optional ByVal fPath As String = vbNullString) As Boolean
    FileExists = (Dir(FormatPath(fName, fPath)) <> vbNullString)
End Function

'Function to check if folder exists
'If path isn't a folder name (can be a file name) then return false
Public Function FolderExists(ByVal fName As String, Optional ByVal fPath As String = vbNullString) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(FormatPath(fName, fPath)) And vbDirectory)
    On Error GoTo 0
End Function

'Function to format file or folder fullPath
Public Function FormatPath(ByVal fName As String, ByVal fPath As String) As String
    'If name starts with "/" or "\" remove first character
    If (Left(fName, 1) = "\" Or Left(fName, 1) = "/") Then
        fName = Right(fName, Len(fName) - 1)
    End If
    
    If (Len(fPath) > 0) Then 'If path isn't empty then concat path and name
        If ((Right(fPath, 1) = "\") Or (Right(fPath, 1) = "/")) Then 'If path ends with "\" or "/" just concat
            fName = fPath & fName
        ElseIf (Len(fPath) > 0) Then 'Else add "\" between path and name
            fName = fPath & "\" & fName
        End If
    End If
    
    'Format to standard windows path
    FormatPath = Replace(fName, "/", "\")
    
    'Remove last "\" if necessary
    If (Right(FormatPath, 1) = "\") Then
        FormatPath = Left(FormatPath, Len(FormatPath) - 1)
    End If
End Function

'Function to get folder root from file name (full path or just file name)
'Assuming the path is full path, if not then root folder is current folder
Public Function GetFolderName(ByVal fPath As String) As String
    GetFolderName = IIf(InStr(fPath, "\") > 0, Left(fPath, InStrRev(fPath, "\") - 1), ThisWorkbook.Path)
End Function

'Function to get file name from path (or last folder if path is a folder path)
Public Function GetSourceName(ByVal fPath As String) As String
    GetSourceName = Right(fPath, Len(fPath) - InStrRev(fPath, "\"))
End Function

'Sub to move or rename file
Public Sub MoveOrRenameFile(ByVal fName As String, ByVal destName As String, Optional ByVal sourcePath As String = vbNullString, Optional ByVal destpath As String = vbNullString)
    fName = FormatPath(fName, sourcePath)
    destName = FormatPath(destName, destpath)
    
    If FileExists(fName) Then
        If (Not FileExists(destName)) Then
            CreateFolder GetFolderName(destName)
            Name fName As destName
        Else
            MsgBox "Destination file already exists.", vbOKOnly + vbExclamation, "Cannot rename file"
        End If
    Else
        MsgBox "Source file not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot rename file"
    End If
End Sub

'Sub to move or rename folder
Public Sub MoveOrRenameFolder(ByVal fName As String, ByVal destName As String, Optional ByVal sourcePath As String = vbNullString, Optional ByVal destpath As String = vbNullString)
    fName = FormatPath(fName, sourcePath)
    destName = FormatPath(destName, destpath)
    
    If FolderExists(fName) Then
        If (Not FolderExists(destName)) Then
            CreateFolder GetFolderName(destName)
            Name fName As destName
        Else
            MsgBox "Destination folder already exists.", vbOKOnly + vbExclamation, "Cannot rename folder"
        End If
    Else
        MsgBox "Source folder not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot rename folder"
    End If
End Sub
