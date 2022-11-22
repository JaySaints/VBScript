Dim objFSO, objFSOText, objFolder, objFile
Dim strDirectory, strFile

const ForReading = 1
Const ForWriting = 2

' Path of directory
strDirectory = "C:\Users\jshark\Documents\Dev\WinCommands\VBScript"

' Name new file
strFileName = "\textFile.txt"

' Text to write in file
strText = "Hello JaySaints!!!"

' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Open and read content file
Set objFile = objFSO.OpenTextFile(strDirectory & strFileName, ForReading)
strOldText = objFile.ReadAll
objFile.Close

' Overwirte file
strNewText = Replace(strText, strOldText, strNewText)

' Open file to Writing
Set objFile = objFSO.OpenTextFile(strDirectory & strFileName, ForWriting)
objFile.WriteLine strNewText
objFile.Close

Wscript.Echo "Edited " & strDirectory & strFileName

Wscript.Quit

