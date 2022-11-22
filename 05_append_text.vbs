Dim objFSO, objFolder, objFile
Dim strDirectory, strFileName

const ForReading = 1
Const ForWriting = 2
Const ForAppening = 8

' Path of directory
strDirectory = "C:\Users\jshark\Documents\Dev\WinCommands\VBScript"

' Name new file
strFileName = "\textFile.txt"

' Text to write in file
strText = "Hello JaySaints!!!"

' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Open file to append content
Set objFile = objFSO.OpenTextFile(strDirectory & strFileName, ForAppening, True)
objFile.WriteLine "User: " & strText
objFile.WriteLine "Age: " & "27"
objFile.Close

Wscript.Echo "Appended text in: " & strDirectory & strFileName

Wscript.Quit

