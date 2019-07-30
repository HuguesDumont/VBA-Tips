Attribute VB_Name = "FileAndFolderFunctions"
Attribute VB_Description = "Functions to handle file and folder processing\r\nNeed ""Microsoft Scripting Runtime"" reference\r\nNeed ""OLE Automation"" reference"
'Need "Microsoft Scripting Runtime" reference
'Need "OLE Automation" reference
Option Explicit

'Function to format file or folder fullPath
Public Function FormatPath(ByVal fName As String, ByVal fPath As String) As String
Attribute FormatPath.VB_Description = "Function to format file or folder fullPath"
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
    FormatPath = Replace(fName, "/", "\")
    
    'Remove last "\" if necessary
    If (Right(FormatPath, 1) = "\") Then
        FormatPath = Left(FormatPath, Len(FormatPath) - 1)
    End If
End Function

'Function to get folder root from file name (full path or just file name)
'Assuming the path is full path, if not then root folder is current folder
Public Function GetFolderName(ByVal fPath As String) As String
Attribute GetFolderName.VB_Description = "Function to get folder root from file name (full path or just file name)\r\nAssuming the path is full path, if not then root folder is current folder"
    If InStr(fPath, "\") > 0 Then
       GetFolderName = Left(fPath, InStrRev(fPath, "\") - 1)
    Else
        GetFolderName = ThisWorkbook.Path
    End If
End Function

'Function to get file name from path (or last folder if path is a folder path)
Public Function GetSourceName(ByVal fPath As String) As String
Attribute GetSourceName.VB_Description = "Function to get file name from path (or last folder if path is a folder path)"
    GetSourceName = Right(fPath, Len(fPath) - InStrRev(fPath, "\"))
End Function

'Function to check if file or folder exists
Public Function FExists(ByVal fName As String, Optional ByVal fPath As String = "") As Boolean
Attribute FExists.VB_Description = "Function to check if file or folder exists"
    FExists = (Len(Dir(FormatPath(fName, fPath), vbDirectory)) <> 0)
End Function

'Function to check if folder exists
'If path isn't a folder name (can be a file name) then return false
Public Function FolderExists(ByVal fName As String, Optional ByVal fPath As String = "") As Boolean
Attribute FolderExists.VB_Description = "Function to check if folder exists\r\nIf path isn't a folder name (can be a file name) then return false"
    On Error Resume Next
    FolderExists = (GetAttr(FormatPath(fName, fPath)) And vbDirectory)
End Function

'Function to check if file exists
Public Function FileExists(ByVal fName As String, Optional ByVal fPath As String = "") As Boolean
Attribute FileExists.VB_Description = "Function to check if file exists"
    FileExists = (Dir(FormatPath(fName, fPath)) <> "")
End Function

'Sub to create folder (if folder doesn't exist)
'If dest = "" then fName is full path
Public Sub CreateFolder(ByVal fName As String, Optional ByVal dest As String = "")
Attribute CreateFolder.VB_Description = "Sub to create folder (if folder doesn't exist)\r\nIf dest = """" then fName is full path"
    On Error GoTo cantCreate
    
    fName = FormatPath(fName, dest)
    
    If Not FolderExists(fName) Then
        Call CreateFolder(GetFolderName(fName))
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
Public Sub CreateFile(ByVal fName As String, Optional ByVal dest As String = "", Optional ByVal overwrite As Boolean = False)
Attribute CreateFile.VB_Description = "Sub to create file (using name and path separately)\r\nIf dest = """" then fName is full path\r\nDo not overwrite by default"
    On Error GoTo cantCreate
    
    Dim FSO As New Scripting.FileSystemObject
    Dim oFile As Object
    
    'Get full path
    fName = FormatPath(fName, dest)
    
    'Create folder if necessary
    CreateFolder GetFolderName(fName)
    
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
Public Sub MoveOrRenameFolder(ByVal fName As String, ByVal destName As String, Optional ByVal sourcePath As String = "", Optional ByVal destPath As String = "")
Attribute MoveOrRenameFolder.VB_Description = "Sub to move or rename folder"
    fName = FormatPath(fName, sourcePath)
    destName = FormatPath(destName, destPath)
    
    If FolderExists(fName) Then
        If Not FolderExists(destName) Then
            CreateFolder GetFolderName(destName)
            Name fName As destName
        Else
            MsgBox "Destination folder already exists.", vbOKOnly + vbExclamation, "Cannot rename folder"
        End If
    Else
        MsgBox "Source folder not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot rename folder"
    End If
End Sub

'Sub to move or rename file
Public Sub MoveOrRenameFile(ByVal fName As String, ByVal destName As String, Optional ByVal sourcePath As String = "", Optional ByVal destPath As String = "")
Attribute MoveOrRenameFile.VB_Description = "Sub to move or rename file"
    fName = FormatPath(fName, sourcePath)
    destName = FormatPath(destName, destPath)
    
    If FileExists(fName) Then
            
        If Not FileExists(destName) Then
            CreateFolder GetFolderName(destName)
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
Public Sub CopyTo(ByVal fName As String, ByVal destPath As String, Optional ByVal sourcePath As String = "", Optional ByVal destName As String = "", Optional ByVal overwrite As Boolean = False)
Attribute CopyTo.VB_Description = "Sub to copy file or folder to another folder\r\nIf dest folder doesn't exists then creates it\r\nIf destName <> """" then copy will have a different name from source"
    If FolderExists(FormatPath(fName, sourcePath)) Then
        Call CopyFolderTo(fName, destPath, sourcePath, destName, copusubfolders, copySubFiles, overwrite)
    ElseIf FileExists(FormatPath(fName, sourcePath)) Then
        Call CopyFileTo(fName, destPath, sourcePath, destName, overwrite)
    Else
        MsgBox "Source not found. Please check your source path.", vbOKOnly + vbExclamation, "Cannot copy"
    End If
End Sub

'Sub to copy a folder to another folder
'If dest folder doesn't exists then creates it
'If destName <> "" then copy will have a different name from source
Public Sub CopyFolderTo(ByVal fName As String, ByVal destPath As String, Optional ByVal sourcePath As String = "", Optional ByVal destName As String = "", Optional ByVal overwrite As Boolean = False)
Attribute CopyFolderTo.VB_Description = "Sub to copy a folder to another folder\r\nIf dest folder doesn't exists then creates it\r\nIf destName <> """" then copy will have a different name from source"
    Dim FSO As New Scripting.FileSystemObject
    
    fName = FormatPath(fName, sourcePath)
    
    If FolderExists(fName) Then
        If destName <> "" Then
            destPath = FormatPath(destName, destPath)
        Else
            destPath = FormatPath(GetSourceName(fName), destPath)
        End If
        
        Call FSO.CopyFolder(fName, destPath, (overwrite Or Not FolderExists(destPath)))
    Else
        MsgBox "Source folder not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot copy folder"
    End If
End Sub

'Sub to copy a file to another folder
'If dest folder doesn't exists then creates it
'If destName <> "" then copy will have a different name from source
Public Sub CopyFileTo(ByVal fName As String, ByVal destPath As String, Optional ByVal sourcePath As String = "", Optional ByVal destName As String = "", Optional ByVal overwrite As Boolean = False)
Attribute CopyFileTo.VB_Description = "Sub to copy a file to another folder\r\nIf dest folder doesn't exists then creates it\r\nIf destName <> """" then copy will have a different name from source"
    fName = FormatPath(fName, sourcePath)
    If FileExists(fName) Then
        If destName <> "" Then
            destPath = FormatPath(destName, destPath)
        Else
            destPath = FormatPath(GetSourceName(fName), destPath)
        End If
        
        If overwrite Or Not FileExists(destPath) Then
            CreateFolder GetFolderName(destPath)
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
Public Sub CopyAllFilesTo(ByVal fName As String, ByVal destPath As String, Optional ByVal sourcePath As String = "", Optional ByVal destName As String = "", Optional ByVal overwrite As Boolean = False)
Attribute CopyAllFilesTo.VB_Description = "Sub to copy all files from one directory to another\r\nIf dest folder doesn't exists then creates it\r\nIf destName <> """" then copy will have a different name from source"
    Dim FSO As New Scripting.FileSystemObject
    
    fName = FormatPath(fName, sourcePath)
    
    If FolderExists(fName) Then
        destPath = FormatPath(destName, destPath)
        CreateFolder destPath
        Call FSO.copyFile(fName & "\*.*", destPath, (overwrite Or Not FolderExists(destPath)))
    Else
        MsgBox "Source folder not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot copy folder"
    End If
End Sub

'Sub to delete folder or file
'If fPath = "" then fName is full Path
Public Sub Deletion(ByVal fName As String, Optional ByVal fPath As String = "")
Attribute Deletion.VB_Description = "Sub to delete folder or file\r\nIf fPath = """" then fName is full Path"
    On Error Resume Next
    deleteFolder fName, fPath
    DeleteFile fName, fPath
End Sub

'Sub to delete folder
'If fPath = "" then fName is full Path
Public Sub deleteFolder(ByVal fName As String, Optional ByVal fPath As String = "")
Attribute deleteFolder.VB_Description = "Sub to delete folder\r\nIf fPath = """" then fName is full Path"
    Dim FSO As New Scripting.FileSystemObject
    
    fName = FormatPath(fName, fPath)
    
    If FolderExists(fName) Then
        On Error Resume Next
        
        'Delete files
        FSO.DeleteFile fName & "\*.*", True
        
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
Public Sub DeleteFile(ByVal fName As String, Optional ByVal fPath As String = "")
Attribute DeleteFile.VB_Description = "Sub to delete file\r\nIf fPath = """" then fName is full Path"
    On Error GoTo errorDeleting
    If FileExists(FormatPath(fName, fPath)) Then
        Kill FormatPath(fName, fPath)
    Else
        MsgBox "Source file not found. Please check your source file path.", vbOKOnly + vbExclamation, "Cannot delete file"
    End If
errorDeleting:
    MsgBox "Error while deleting file. Please check that your file is closed and you have the correct rigths.", vbOKOnly + vbCritical, "Cannot delete file"
End Sub

'Sub to delete all files in folder
'If fPath = "" then fName is full path
Public Sub DeleteAllFiles(ByVal fName As String, Optional ByVal fPath As String = "")
Attribute DeleteAllFiles.VB_Description = "Sub to delete all files in folder\r\nIf fPath = """" then fName is full path"
    On Error GoTo errorDeleting
    
    If FolderExists(FormatPath(fName, fPath)) Then
        Kill FormatPath(fName, fPath) & "*.*"
    Else
        MsgBox "Source folder not found. Please check your source folder path.", vbOKOnly + vbExclamation, "Cannot delete files"
    End If
    
errorDeleting:
    MsgBox "Error while deleting files. Please check that your files are closed and you have the correct rigths.", vbOKOnly + vbCritical, "Cannot delete files"
End Sub
