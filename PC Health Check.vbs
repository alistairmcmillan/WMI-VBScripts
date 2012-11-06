' Author: Alistair McMillan
' Start Date: 09 May 2011
' -----------------------------------------------

Option Explicit

Dim freePhysicalMemory, objSWbemLocator, objSWbemServices, objItem, colOSItems, _
	colItems, strComputer, strResult, freeSpaceInPagingFiles, sizeStoredInPagingFiles

strComputer = InputBox("Enter full computer name (i.e. SWSA29565) or IP address")

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = GetObject( "winmgmts://" & strComputer & "/root/cimv2" )
objSWbemServices.Security_.ImpersonationLevel = 3

' Basic OS info
Set colOSItems = objSWbemServices.ExecQuery("select * from Win32_OperatingSystem")
For Each objItem in colOSItems
	strResult = strResult & "Operating System: " & objItem.Caption & vbCr
	freePhysicalMemory = objItem.FreePhysicalMemory
	freeSpaceInPagingFiles = objItem.FreeSpaceInPagingFiles
	sizeStoredInPagingFiles = objItem.SizeStoredInPagingFiles
Next

' Basic info
Set colItems = objSWbemServices.ExecQuery("select * from Win32_ComputerSystem")
For Each objItem in colItems
	strResult = strResult & "Domain: " & objItem.Domain & vbCr
	strResult = strResult & "SystemName: " & objItem.Name & vbCr
	strResult = strResult & "Current user: " & objItem.UserName & vbCr
	strResult = strResult & vbCr
	strResult = strResult & "Free RAM: " & Round(freePhysicalMemory/1024, 1) & " MB" & vbCr
	strResult = strResult & "Total RAM: " & Round(objItem.TotalPhysicalMemory/1024/1024, 1) & " MB" & vbCr
	strResult = strResult & vbCr
	strResult = strResult & "Free Page Files Size: " & Round(freeSpaceInPagingFiles/1024, 1) & " MB" & vbCr
	strResult = strResult & "Total Page Files Size: " & Round(sizeStoredInPagingFiles/1024, 1) & " MB" & vbCr
Next

' RAM

' Lists disks with total/free space
Set colItems = objSWbemServices.ExecQuery("select * from Win32_LogicalDisk")
For Each objItem in colItems
	If ("C:" = objItem.Name) Then
		strResult = strResult & vbCr & objItem.Name & VbCr
		If (IsNull(objItem.Size) Or IsNull(objItem.FreeSpace)) Then
			strResult = strResult & "Free Space: " & objItem.FreeSpace/1024/1024/1024 & " GB" & VbCr
			strResult = strResult & "Total Space: " & objItem.Size/1024/1024/1024 & " GB" & VbCr
			strResult = strResult & "Percent free: " & objItem.FreeSpace/objItem.Size & " %" & VbCr
		else
			strResult = strResult & "Free Space: " & Round(objItem.FreeSpace/1024/1024/1024, 2) & " GB" & vbCr
			strResult = strResult & "Total Space: " & Round(objItem.Size/1024/1024/1024, 2) & " GB" & VbCr
			strResult = strResult & "Percent free: " & Round((objItem.FreeSpace/objItem.Size)*100, 1) & "%" & vbCr
		End If
	End If
Next

Wscript.Echo strResult

WSCript.Quit
