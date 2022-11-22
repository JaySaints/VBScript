Dim strDirectory, strFile

strDirectory = "C:\Users\jshark\Documents\Dev\WinCommands\jota"
strFile = "\Summer.txt"


Function CreateNewFile(strFolderName, strFileName)
    
    Dim objFSO, objFile

    ' Create the File System Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If Not objFSO.FileExists(strFolderName & strFileName) Then
    
        If Not objFSO.FolderExists(strFolderName) Then
            ' Create the Folder specified by strFolderName
            Set objFolder = objFSO.CreateFolder(strFolderName)
            
        End If

        'Creates the file using the value of strFile 
        Set objFile = objFSO.CreateTextFile(strFolderName & strFileName)

    End If

End Function ' CreateNewFile


Call CreateNewFile(strDirectory, strFile)

Wscript.Echo "Just created " & strDirectory & strFile

Wscript.Quit
