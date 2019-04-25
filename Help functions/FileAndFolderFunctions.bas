Attribute VB_Name = "FileAndFolderFunctions"
Attribute VB_Description = "Functions to handle file and folder processing\r\nNeed ""Microsoft Scripting Runtime"" reference\r\nNeed ""OLE Automation"" reference"
'Need "Microsoft Scripting Runtime" reference
'Need "OLE Automation" reference
Option Explicit

'Function to format file or folder fullPath
Public Function formatPath(ByVal fName As String, ByVal fPath As String) As String
Attribute formatPath.VB_Description = "Function to format file or folder fullPath"
    'If name starts with "/" or "\" remove first character
    If (Left(fName, 1) = "\" Or Left(fName, 1) = "/") Then
        fName = Right(fName, Len(fName) - 1)
    End If
    'If path isn't empty then concat path and name
    If fPath <> "" Then
        'If path ends with "\" or "/" just concat
        If (Right(fPath, 1) = "\" Or Right(fPath, 1) = "/") Then
            fName = fPath & fName
        'Else had "\" between path and name
        ElseIf (fPath <> "") Then
            fName = fPath & "\" & fName
        End If
    End If
    
    'Format to standard windows path
    formatPath = Replace(fName, "/", "\")
    
    'Remove last "\" if necessary
    If (Right(formatPath, 1) = "\") Then
        formatPath = Left(formatPath, Len(formatPath) - 1)
    End If
End Function

'Function to get folder root from file name (full path or just file name)
'Assuming the path is full path, if not then root folder is current folder
Public Function getFolderName(fPath As String) As String
Attribute getFolderName.VB_Description = "Function to get folder root from file name (full path or just file name)\r\nAssuming the path is full path, if not then root folder is current folder"
    If InStr(fPath, "\") > 0 Then
       getFolderName = Left(fPath, InStrRev(fPath, "\") - 1)
    Else
        getFolderName = ThisWorkbook.Path
    End If
End Function

'Function to get file name from path (or last folder if path is a folder path)
Public Function getSourceName(fPath As String) As String
Attribute getSourceName.VB_Description = "Function to get file name from path (or last folder if path is a folder path)"
    getSourceName = Right(fPath, Len(fPath) - InStrRev(fPath, "\"))
End Function

'Function to check if file or folder exists
Public Function fExists(fName As String, Optional fPath As String = "") As Boolean
Attribute fExists.VB_Description = "Function to check if file or folder exists"
    fExists = (Len(Dir(formatPath(fName, fPath), vbDirectory)) <> 0)
End Function

'Function to check if folder exists
'If path isn't a folder name (can be a file name) then return false
Public Function folderExists(fName As String, Optional fPath As String = "") As Boolean
Attribute folderExists.VB_Description = "Function to check if folder exists\r\nIf path isn't a folder name (can be a file name) then return false"
    On Error Resume Next
    folderExists = (GetAttr(formatPath(fName, fPath)) And vbDirectory)
End Function

'Function to check if file exists
Public Function fileExists(fName As String, Optional fPath As String = "") As Boolean
Attribute fileExists.VB_Description = "Function to check if file exists"
    fileExists = (Dir(formatPath(fName, fPath)) <> "")
End Function

'Sub to create folder (if folder doesn't exist)
'If dest = "" then fName is full path
Public Sub createFolder(fName As String, Optional dest As String = "")
Attribute createFolder.VB_Description = "Sub to create folder (if folder doesn't exist)\r\nIf dest = """" then fName is full path"
    On Error GoTo cantCreate
    
    fName = formatPath(fName, dest)
    
    If Not folderExists(fName) Then
        Call createFolder(getFolderName(fName))
        MkDir fName
    End If
    
    Exit Sub
cantCreate:
    MsgBox "Error while trying to create new folder. Please check your path and folder name don't contain incompatible characters or if you have the rights to create the folder.", _
        vbOKOnly + vbCritical, "Error creating folder"
End Sub

'Sub to create file (using name and path separately)
'If dest = "" then fName is full path
'Do not overwrite by default
Public Sub createFile(fName As String, Optional dest As String = "", Optional overwrite As Boolean = False)
Attribute createFile.VB_Description = "Sub to create file (using name and path separately)\r\nIf dest = """" then fName is full path\r\nDo not overwrite by default"
    On Error GoTo cantCreate
    
    Dim FSO As New Scripting.FileSystemObject
    Dim oFile As Object
    
    'Get full path
    fName = formatPath(fName, dest)
    
    'Create folder if necessary
    createFolder getFolderName(fName)
    
    'Create the file
    Set oFile = FSO.createTextFile(fName, overwrite)
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

'Sub to move or rename folder
Public Sub moveOrRenameFolder(fName As String, destName As String, Optional sourcePath As String = "", Optional destPath As String = "")
Attribute moveOrRenameFolder.VB_Description = "Sub to move or rename folder"
    fName = formatPath(fName, sourcePath)
    destName = formatPath(destName, destPath)
    
    If folderExists(fName) Then
        If Not folderExists(destName) Then
            createFolder getFolderName(destName)
            Name fName As destName
        Else
            MsgBox "Destination folder already exists.", vbOKOnly + vbExclamation, "Cannot rename folder"
        End If
    Else
        MsgBox "Source folder not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot rename folder"
    End If
End Sub

'Sub to move or rename file
Public Sub moveOrRenameFile(fName As String, destName As String, Optional sourcePath As String = "", Optional destPath As String = "")
Attribute moveOrRenameFile.VB_Description = "Sub to move or rename file"
    fName = formatPath(fName, sourcePath)
    destName = formatPath(destName, destPath)
    
    If fileExists(fName) Then
            
        If Not fileExists(destName) Then
            createFolder getFolderName(destName)
            Name fName As destName
        Else
            MsgBox "Destination file already exists.", vbOKOnly + vbExclamation, "Cannot rename file"
        End If
    Else
        MsgBox "Source file not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot rename file"
    End If
End Sub

'Sub to copy file or folder to another folder
'If dest folder doesn't exists then creates it
'If destName <> "" then copy will have a different name from source
Public Sub copyTo(fName As String, destPath As String, Optional sourcePath As String = "", Optional destName As String = "", Optional overwrite As Boolean = False)
Attribute copyTo.VB_Description = "Sub to copy file or folder to another folder\r\nIf dest folder doesn't exists then creates it\r\nIf destName <> """" then copy will have a different name from source"
    
    If folderExists(formatPath(fName, sourcePath)) Then
        Call copyFolderTo(fName, destPath, sourcePath, destName, copusubfolders, copySubFiles, overwrite)
    ElseIf fileExists(formatPath(fName, sourcePath)) Then
        Call copyFileTo(fName, destPath, sourcePath, destName, overwrite)
    Else
        MsgBox "Source not found. Please check your source path.", vbOKOnly + vbExclamation, "Cannot copy"
    End If
End Sub

'Sub to copy a folder to another folder
'If dest folder doesn't exists then creates it
'If destName <> "" then copy will have a different name from source
Public Sub copyFolderTo(fName As String, destPath As String, Optional sourcePath As String = "", Optional destName As String = "", Optional overwrite As Boolean = False)
Attribute copyFolderTo.VB_Description = "Sub to copy a folder to another folder\r\nIf dest folder doesn't exists then creates it\r\nIf destName <> """" then copy will have a different name from source"
    Dim FSO As New Scripting.FileSystemObject
    
    fName = formatPath(fName, sourcePath)
    
    If folderExists(fName) Then
        If destName <> "" Then
            destPath = formatPath(destName, destPath)
        Else
            destPath = formatPath(getSourceName(fName), destPath)
        End If
        
        Call FSO.CopyFolder(fName, destPath, (overwrite Or Not folderExists(destPath)))
    Else
        MsgBox "Source folder not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot copy folder"
    End If
End Sub

'Sub to copy a file to another folder
'If dest folder doesn't exists then creates it
'If destName <> "" then copy will have a different name from source
Public Sub copyFileTo(fName As String, destPath As String, Optional sourcePath As String = "", Optional destName As String = "", Optional overwrite As Boolean = False)
Attribute copyFileTo.VB_Description = "Sub to copy a file to another folder\r\nIf dest folder doesn't exists then creates it\r\nIf destName <> """" then copy will have a different name from source"
    fName = formatPath(fName, sourcePath)
    If fileExists(fName) Then
        If destName <> "" Then
            destPath = formatPath(destName, destPath)
        Else
            destPath = formatPath(getSourceName(fName), destPath)
        End If
        
        If overwrite Or Not fileExists(destPath) Then
            createFolder getFolderName(destPath)
            FileCopy fName, destPath
        Else
            MsgBox "File already exists in destination folder.", vbOKOnly + vbExclamation, "File already exists"
        End If
    Else
        MsgBox "Source file not found. Please check your source file path.", vbOKOnly + vbExclamation, "Cannot copy file"
    End If
End Sub

'Sub to copy all files from one directory to another
'If dest folder doesn't exists then creates it
'If destName <> "" then copy will have a different name from source
Public Sub copyAllFilesTo(fName As String, destPath As String, Optional sourcePath As String = "", Optional destName As String = "", Optional overwrite As Boolean = False)
Attribute copyAllFilesTo.VB_Description = "Sub to copy all files from one directory to another\r\nIf dest folder doesn't exists then creates it\r\nIf destName <> """" then copy will have a different name from source"
    Dim FSO As New Scripting.FileSystemObject
    
    fName = formatPath(fName, sourcePath)
    
    If folderExists(fName) Then
        destPath = formatPath(destName, destPath)
        createFolder destPath
        Call FSO.copyFile(fName & "\*.*", destPath, (overwrite Or Not folderExists(destPath)))
    Else
        MsgBox "Source folder not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot copy folder"
    End If
End Sub

'Sub to delete folder or file
'If fPath = "" then fName is full Path
Public Sub deletion(fName As String, Optional fPath As String = "")
Attribute deletion.VB_Description = "Sub to delete folder or file\r\nIf fPath = """" then fName is full Path"
    On Error Resume Next
    deleteFolder fName, fPath
    deleteFile fName, fPath
End Sub

'Sub to delete folder
'If fPath = "" then fName is full Path
Public Sub deleteFolder(fName As String, Optional fPath As String = "")
Attribute deleteFolder.VB_Description = "Sub to delete folder\r\nIf fPath = """" then fName is full Path"
    Dim FSO As New Scripting.FileSystemObject
    
    fName = formatPath(fName, fPath)
    
    If folderExists(fName) Then
        On Error Resume Next
        
        'Delete files
        FSO.deleteFile fName & "\*.*", True
        
        'Delete subfolders
        FSO.deleteFolder fName & "\*.*", True
        
        Set FSO = Nothing
        
        RmDir fName
        
        On Error GoTo 0
    Else
        MsgBox "Source folder not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot delete folder"
    End If
End Sub

'Sub to delete file
'If fPath = "" then fName is full Path
Public Sub deleteFile(fName As String, Optional fPath As String = "")
Attribute deleteFile.VB_Description = "Sub to delete file\r\nIf fPath = """" then fName is full Path"
    On Error GoTo errorDeleting
    If fileExists(formatPath(fName, fPath)) Then
        Kill formatPath(fName, fPath)
    Else
        MsgBox "Source file not found. Please check your source file path.", vbOKOnly + vbExclamation, "Cannot delete file"
    End If
errorDeleting:
    MsgBox "Error while deleting file. Please check that your file is closed and you have the correct rigths.", vbOKOnly + vbCritical, "Cannot delete file"
End Sub

'Sub to delete all files in folder
'If fPath = "" then fName is full path
Public Sub deleteAllFiles(fName As String, Optional fPath As String = "")
Attribute deleteAllFiles.VB_Description = "Sub to delete all files in folder\r\nIf fPath = """" then fName is full path"
    On Error GoTo errorDeleting
    
    If folderExists(formatPath(fName, fPath)) Then
        Kill formatPath(fName, fPath) & "*.*"
    Else
        MsgBox "Source folder not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot delete files"
    End If
    
errorDeleting:
    MsgBox "Error while deleting files. Please check that your files are closed and you have the correct rigths.", vbOKOnly + vbCritical, "Cannot delete files"
End Sub
