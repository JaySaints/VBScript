
' Path of directory
strDirectory = "C:\Users\jshark\Documents\Dev\WinCommands\VBScript"

' Name new file
strFileName = "\names_info.txt"

' Function 
Function appendText(pathFile, strContext)
Dim objFSO, objFolder, objFile
Dim strDirectory, strFileName

Const ForAppening = 8
    
' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Open file to append content
Set objFile = objFSO.OpenTextFile(pathFile, ForAppening, True)
objFile.WriteLine strContext
objFile.Close

End Function ' appendText





strComputer = "."
Set objSWbemServices = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colSWbemObjectSet = objSWbemServices.ExecQuery _
   ("SELECT * FROM Win32_Service")
For Each objSWbemObject In colSWbemObjectSet
    Call appendText(strDirectory & strFileName, "Name: " & objSWbemObject.Name)
Next