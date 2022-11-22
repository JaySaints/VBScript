' Author: Pablo J. Santos
' Created: 22 November 2022

' Path of directory
strDirectory = "C:\Users\jshark\Documents\Dev\WinCommands\VBScript"

' Name new file
strFileName = "\system_informations.txt"

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

' [WMI - Windows Management Instrumentation]
' For more information access: (https://learn.microsoft.com/en-us/windows/win32/wmisdk/about-wmi) 
strComputer = "."
Set objSWbemServices = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colSWbemObjectSet = objSWbemServices.ExecQuery("SELECT * FROM Win32_OperatingSystem")
Set colNetAdapConfigItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration")
Set colDiskItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_LogicalDisk") 
Set objSWbemDateTime = CreateObject("WbemScripting.SWbemDateTime")

' Get Information about OPERATION SYSTEM
For Each objSWbemObject In colSWbemObjectSet
    Call appendText(strDirectory & strFileName, "Hostname: " & objSWbemObject.CSName)
    Call appendText(strDirectory & strFileName, "OS: " & objSWbemObject.Caption)
    Call appendText(strDirectory & strFileName, "Version: " & objSWbemObject.Version)
    Call appendText(strDirectory & strFileName, "SerialNumber: " & objSWbemObject.SerialNumber)
    Call appendText(strDirectory & strFileName, "Status: " & objSWbemObject.Status)

    objSWbemDateTime.Value = objSWbemObject.InstallDate
    Call appendText(strDirectory & strFileName, "Install Date: " & objSWbemDateTime.GetVarDate(True))

    objSWbemDateTime.Value = objSWbemObject.LastBootUpTime
    Call appendText(strDirectory & strFileName, "Last Boot UpTime: " & objSWbemDateTime.GetVarDate(True))
Next

' Get Information about NETWORKING CONFIGURATION
For Each objNetAdapConfigItem in colNetAdapConfigItems
    If isNull(objNetAdapConfigItem.IPAddress) Then
        '// Skip adapter, not currently used
    Else
        Call appendText(strDirectory & strFileName, "Network Adapter: " & objNetAdapConfigItem.Description)
        Call appendText(strDirectory & strFileName, "MAC Address: " & objNetAdapConfigItem.MACAddress)
        Call appendText(strDirectory & strFileName, "DHCP Enabled: " & objNetAdapConfigItem.DHCPEnabled)
        Call appendText(strDirectory & strFileName, "IP Address: " & Join(objNetAdapConfigItem.IPAddress, ","))
        Call appendText(strDirectory & strFileName, "Subnet Mask: " & Join(objNetAdapConfigItem.IPSubnet, ","))
    End If
Next

' Get Information about LOGICAL DISK
For Each objDiskItem in colDiskItems
    If objDiskItem.DriveType = 3 Then
        Call appendText(strDirectory & strFileName, "Volume: " & objDiskItem.Caption)
        Call appendText(strDirectory & strFileName, "Compressed: " & objDiskItem.Compressed)
        Call appendText(strDirectory & strFileName, "File System: " & objDiskItem.FileSystem)
        Call appendText(strDirectory & strFileName, "Volume Size: " & FormatNumber(objDiskItem.Size/1024^3, 2) & " GB")
        Call appendText(strDirectory & strFileName, "Free Space: " & FormatNumber(objDiskItem.FreeSpace/1024^3, 2) & " GB")
    End If
Next




WScript.Echo("Clonclued!")