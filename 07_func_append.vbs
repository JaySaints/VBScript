Dim objFSO, objFolder, objFile
Dim strDirectory, strFileName

Const ForAppening = 8

Function appendText(pathFile, strContext)
    
' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Open file to append content
Set objFile = objFSO.OpenTextFile(pathFile, ForAppening, True)
objFile.WriteLine strContext
objFile.Close

End Function ' appendText

' Path of directory
strDirectory = "C:\Users\jshark\Documents\Dev\WinCommands\VBScript"

' Name new file
strFileName = "\textFile.txt"

' Text to write in file
strText = "It's Worked!!!"

'Call function

Call appendText(strDirectory & strFileName, strText)

Wscript.Echo "Appended text in: " & strDirectory & strFileName

Wscript.Quit

