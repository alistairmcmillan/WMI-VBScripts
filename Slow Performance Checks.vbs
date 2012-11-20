' Author: Alistair McMillan
' Start Date: 12 November 2012
' -----------------------------------------------

Option Explicit

Const HKEY_LOCAL_MACHINE = &H80000002 'HKEY_LOCAL_MACHINE

Dim freePhysicalMemory, objSWbemLocator, objSWbemServices, objItem, colOSItems, _
	colItems, strComputer, WshShell, freeSpaceInPagingFiles, sizeStoredInPagingFiles, _
	strFilename, objFileSystem, objFile, strMachineName, objFSO, _
	objFolder, objSubFolders, objSubFolder, queriesFolder, objRegistry, _
	values, strValue, strValues, strKeyPath, strValueName, SchTasksCommand, _
	oExec, Line, servicePackeMajor, servicePackMinor

Set WshShell = CreateObject("WScript.Shell")

strComputer = InputBox("Enter full computer name (i.e. SWSA29565) or IP address")

If (IsNull(strComputer) Or IsEmpty(strComputer) Or Len(strComputer) < 1) Then

	Wscript.Echo "Can't continue without a machine name or IP address."

Else

	Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
	Set objSWbemServices = GetObject( "winmgmts://" & strComputer & "/root/cimv2" )
	objSWbemServices.Security_.ImpersonationLevel = 3

	' Get machine name to user for filename
	Set colItems = objSWbemServices.ExecQuery("select * from Win32_ComputerSystem")
	For Each objItem in colItems
		strMachineName = objItem.Name
	Next

	strFilename = strMachineName + " slow performance checks.txt"
	Set objFileSystem = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFileSystem.OpenTextFile(strFilename, 8, True)

	objFile.WriteLine("")
	objFile.WriteLine("==================================================")
	objFile.WriteLine("==================================================")
	objFile.WriteLine("")
	objFile.WriteLine("Scanned: " & Now)
	objFile.WriteLine("")

	' Basic memory info
	Set colOSItems = objSWbemServices.ExecQuery("select * from Win32_OperatingSystem")
	For Each objItem in colOSItems
		objFile.WriteLine("Operating System: " & objItem.Caption)
		objFile.WriteLine("Service Pack Major: " & objItem.ServicePackMajorVersion)
		objFile.WriteLine("Service Pack Minor: " & objItem.ServicePackMinorVersion)
		freePhysicalMemory = objItem.FreePhysicalMemory
		freeSpaceInPagingFiles = objItem.FreeSpaceInPagingFiles
		sizeStoredInPagingFiles = objItem.SizeStoredInPagingFiles
	Next

	' Basic info
	Set colItems = objSWbemServices.ExecQuery("select * from Win32_ComputerSystem")
	For Each objItem in colItems
		objFile.WriteLine("Domain: " & objItem.Domain)
		objFile.WriteLine("SystemName: " & objItem.Name)
		objFile.WriteLine("Current user: " & objItem.UserName)
		objFile.WriteLine("")
		objFile.WriteLine("Free RAM: " & Round(freePhysicalMemory/1024, 1) & " MB")
		objFile.WriteLine("Total RAM: " & Round(objItem.TotalPhysicalMemory/1024/1024, 1) & " MB")
		objFile.WriteLine("")
		objFile.WriteLine("Free Page Files Size: " & Round(freeSpaceInPagingFiles/1024, 1) & " MB")
		objFile.WriteLine("Total Page Files Size: " & Round(sizeStoredInPagingFiles/1024, 1) & " MB")
	Next

	objFile.WriteLine("")

	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management"
	strValueName = "PagingFiles"
	objRegistry.GetMultiStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, values
	For Each strValue in values
		If (IsNull(strValue)) Then
			objFile.WriteLine("No paging files key")
		Else
			If (InStr(strValue, " 0 0")) Then
				objFile.WriteLine("Virtual memory is SYSTEM MANAGED - " & strValue)
			Else
				objFile.WriteLine("Virtual memory is NOT system managed - " & strValue)
			End If
		End If
	Next

	' CPU
	Set colItems = objSWbemServices.ExecQuery("Select * from Win32_Processor")
	For Each objItem in colItems
		objFile.WriteLine("")
		objFile.WriteLine("Device ID: " & objItem.DeviceID)
		objFile.WriteLine("Current Clock Speed: " & objItem.CurrentClockSpeed)
		objFile.WriteLine("Maximum Clock Speed: " & objItem.MaxClockSpeed)
		objFile.WriteLine("Load Percentage: " & objItem.LoadPercentage)
		objFile.WriteLine("Address Width: " & objItem.AddressWidth)
		objFile.WriteLine("Data Width: " & objItem.DataWidth)
		objFile.WriteLine("Manufacturer: " & objItem.Manufacturer)
		objFile.WriteLine("Name: " & objItem.Name)
		objFile.WriteLine("Description: " & objItem.Description)
	Next

	' Lists disks with total/free space
	Set colItems = objSWbemServices.ExecQuery("select * from Win32_LogicalDisk")
	For Each objItem in colItems
		If ("C:" = objItem.Name) Then
			objFile.WriteLine("")
			objFile.WriteLine(objItem.Name)
			If (IsNull(objItem.Size) Or IsNull(objItem.FreeSpace)) Then
				objFile.WriteLine("Free Space: " & objItem.FreeSpace/1024/1024/1024 & " GB")
				objFile.WriteLine("Total Space: " & objItem.Size/1024/1024/1024 & " GB")
				objFile.WriteLine("Percent free: " & objItem.FreeSpace/objItem.Size & " %")
			else
				objFile.WriteLine("Free Space: " & Round(objItem.FreeSpace/1024/1024/1024, 2) & " GB")
				objFile.WriteLine("Total Space: " & Round(objItem.Size/1024/1024/1024, 2) & " GB")
				objFile.WriteLine("Percent free: " & Round((objItem.FreeSpace/objItem.Size)*100, 1) & "%")
			End If
		End If
	Next

	objFile.WriteLine("")

	objFile.WriteLine("SCHEDULED TASKS")
	objFile.WriteLine("---------------")

	SchTasksCommand = "schtasks /query /v /FO TABLE /s " & strMachineName
	Set oExec = WshShell.Exec(SchTasksCommand)

	On Error Resume Next
	Do While oExec.StdOut.AtEndOfStream <> True
		Line = oExec.StdOut.ReadLine
		If Not IsNull(Line) Then
			objFile.WriteLine(Line)
		End If
	Loop

	objFile.WriteLine("")

	objFile.WriteLine("TEMP FOLDERS")
	objFile.WriteLine("------------")
	objFile.WriteLine("FOLDER SIZE (MB), FOLDER NAME")

	If objFileSystem.FolderExists("\\" & strComputer & "\c$\windows\temp") Then
		Set queriesFolder = objFileSystem.GetFolder("\\" & strComputer & "\c$\windows\temp")
		On Error Resume Next
		objFile.Write(Round(queriesFolder.Size/1024/1024, 1))
		objFile.Write(", <" & queriesFolder.Path & ">")
		objFile.WriteLine(" ")
	End If

	If objFileSystem.FolderExists("\\" & strComputer & "\c$\temp") Then
		Set queriesFolder = objFileSystem.GetFolder("\\" & strComputer & "\c$\temp")
		On Error Resume Next
		objFile.Write(Round(queriesFolder.Size/1024/1024, 1))
		objFile.Write(", <" & queriesFolder.Path & ">")
		objFile.WriteLine(" ")
	End If

	Set objFolder = objFileSystem.GetFolder("\\" & strComputer & "\C$\documents and settings")
	Set objSubFolders = objFolder.SubFolders
	For Each objSubFolder in objSubFolders
		If objFileSystem.FolderExists("\\" & strComputer & "\c$\documents and settings\" & objSubFolder.name & "\local settings\temp") Then
			Set queriesFolder = objFileSystem.GetFolder("\\" & strComputer & "\c$\documents and settings\" & objSubFolder.name & "\local settings\temp")
			On Error Resume Next
			objFile.Write(Round(queriesFolder.Size/1024/1024, 1))
			objFile.Write(", <" & queriesFolder.Path & ">")
			objFile.WriteLine(" ")
		End If
	Next

	objFile.WriteLine("")

	objFile.WriteLine("TEMPORARY INTERNET FOLDERS")
	objFile.WriteLine("--------------------------")
	objFile.WriteLine("FOLDER SIZE (MB), FOLDER NAME")

	Set objFolder = objFileSystem.GetFolder("\\" & strComputer & "\C$\documents and settings")
	Set objSubFolders = objFolder.SubFolders
	For Each objSubFolder in objSubFolders
		If objFileSystem.FolderExists("\\" & strComputer & "\c$\documents and settings\" & objSubFolder.name & "\local settings\temporary internet files") Then
			Set queriesFolder = objFileSystem.GetFolder("\\" & strComputer & "\c$\documents and settings\" & objSubFolder.name & "\local settings\temporary internet files")
			On Error Resume Next
			objFile.Write(Round(queriesFolder.Size/1024/1024, 1))
			objFile.Write(", <" & queriesFolder.Path & ">")
			objFile.WriteLine(" ")
		End If
	Next

	objFile.WriteLine("")

	objFile.WriteLine("QUERIES FOLDERS")
	objFile.WriteLine("---------------")
	objFile.WriteLine("FOLDER SIZE (MB), FOLDER NAME")

	Set objFolder = objFileSystem.GetFolder("\\" & strComputer & "\C$\documents and settings")
	Set objSubFolders = objFolder.SubFolders
	For Each objSubFolder in objSubFolders
		If objFileSystem.FolderExists("\\" & strComputer & "\c$\documents and settings\" & objSubFolder.name & "\application data\microsoft\queries") Then
			Set queriesFolder = objFileSystem.GetFolder("\\" & strComputer & "\c$\documents and settings\" & objSubFolder.name & "\application data\microsoft\queries")
			On Error Resume Next
			objFile.Write(Round(queriesFolder.Size/1024/1024, 1))
			objFile.Write(", <" & queriesFolder.Path & ">")
			objFile.WriteLine(" ")
		End If
	Next

	'Wscript.Echo "Finished writing " + strFilename

	SchTasksCommand = "notepad.exe " & strFilename
	WshShell.Exec(SchTasksCommand)

End If
	
Wscript.Quit
