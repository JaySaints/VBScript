' Author: Pablo J. Santos
' Created: 22 November 2022

Dim objFSO, objFile
Dim strDirectory, strFileName

' Path of directory
strDirectory = "C:\Users\..."

' Name new file
strFileName = "\system_informations.txt"

' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

'If File or Directory not exist it will to create
If Not objFSO.FileExists(strDirectory & strFileName) Then
    If Not objFSO.FolderExists(strDirectory) Then        
        objFSO.CreateFolder(strDirectory)        
    End If
    objFSO.CreateTextFile(strDirectory & strFileName)
End If
  
Function AppendText(strContent)    
    ' Open file to append content >> 8 is the value to append text in file
    Set objFile = objFSO.OpenTextFile(strDirectory & strFileName, 8, True)
    objFile.WriteLine strContent
    objFile.Close
End Function ' AppendText

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
    Call AppendText("Hostname: " & objSWbemObject.CSName)
    Call AppendText("OS: " & objSWbemObject.Caption)
    Call AppendText("Version: " & objSWbemObject.Version)
    Call AppendText("SerialNumber: " & objSWbemObject.SerialNumber)
    Call AppendText("Status: " & objSWbemObject.Status)

    objSWbemDateTime.Value = objSWbemObject.InstallDate
    Call AppendText("Install Date: " & objSWbemDateTime.GetVarDate(True))

    objSWbemDateTime.Value = objSWbemObject.LastBootUpTime
    Call AppendText("Last Boot UpTime: " & objSWbemDateTime.GetVarDate(True))
Next

' Get Information about NETWORKING CONFIGURATION
For Each objNetAdapConfigItem in colNetAdapConfigItems
    If isNull(objNetAdapConfigItem.IPAddress) Then
        '// Skip adapter, not currently used
    Else
        Call AppendText("Network Adapter: " & objNetAdapConfigItem.Description)
        Call AppendText("MAC Address: " & objNetAdapConfigItem.MACAddress)
        Call AppendText("DHCP Enabled: " & objNetAdapConfigItem.DHCPEnabled)
        Call AppendText("IP Address: " & Join(objNetAdapConfigItem.IPAddress, ","))
        Call AppendText("Subnet Mask: " & Join(objNetAdapConfigItem.IPSubnet, ","))
    End If
Next

' Get Information about LOGICAL DISK
For Each objDiskItem in colDiskItems
    If objDiskItem.DriveType = 3 Then
        Call AppendText("Volume: " & objDiskItem.Caption)
        Call AppendText("Compressed: " & objDiskItem.Compressed)
        Call AppendText("File System: " & objDiskItem.FileSystem)
        Call AppendText("Volume Size: " & FormatNumber(objDiskItem.Size/1024^3, 2) & " GB")
        Call AppendText("Free Space: " & FormatNumber(objDiskItem.FreeSpace/1024^3, 2) & " GB")
    End If
Next


WScript.Echo("Clonclued!")
