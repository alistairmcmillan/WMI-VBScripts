' Slow Performance Checks.vbs v2.4
' Author: Alistair McMillan
' Start Date: 12 November 2012
' ----------------------------

Option Explicit

Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CURRENT_USER = &H80000001
const REG_SZ = 1
const REG_EXPAND_SZ = 2
const REG_BINARY = 3
const REG_DWORD = 4
const REG_MULTI_SZ = 7

Dim freePhysicalMemory, objSWbemLocator, objSWbemServices, objItem, colOSItems, _
	colItems, strComputer, WshShell, freeSpaceInPagingFiles, sizeStoredInPagingFiles, _
	strFilename, objFileSystem, objFile, strMachineName, objFSO, _
	objFolder, objSubFolders, objSubFolder, queriesFolder, objRegistry, _
	values, strValue, strValues, strKeyPath, strValueName, SchTasksCommand, _
	oExec, Line, servicePackeMajor, servicePackMinor, arrValueNames, arrValueTypes, _
	index, strOperatingSystem, strServiceParkMajor, strServiceParkMinor, strCurrentUser, _
	hasQueriesProblem, tempFolderTotal, temporaryInternetFolderTotal, queriesFolderTotal, _
	boolRemoteStartupItem, boolTempFoldersProblem, boolTempInternetFoldersProblem, _
	boolServicePackMissing, DataList, strOutput, pathString, objPrinter, colPrinters, _
	strPrinterType, boolVirtMemNotSystemManaged, boolDefaultTempProblem, _
	boolDefaultTempInternetFilesProblem, boolAncientPrinterQueues, boolRemotePathProblem, _
	strIPAddress, strLastBootUpTime, boolNeedsReboot, boolWindows2000

Function PadNumbers(input)
	Dim output
	If (input < 10000) Then
		output = output & " "
	End If
	If (input < 1000) Then
		output = output & "  "
	End If
	If (input < 100) Then
		output = output & " "
	End If
	If (input < 10) Then
		output = output & " "
	End If
	PadNumbers = output
End Function

' From http://www.activexperts.com/admin/scripts/vbscript/0360/
Function WMIDateStringToDate(dtmBootup)
    WMIDateStringToDate = CDate(Mid(dtmBootup, 5, 2) & "/" & _
        Mid(dtmBootup, 7, 2) & "/" & Left(dtmBootup, 4) _
            & " " & Mid (dtmBootup, 9, 2) & ":" & _
                Mid(dtmBootup, 11, 2) & ":" & Mid(dtmBootup,13, 2))
End Function

Set WshShell = CreateObject("WScript.Shell")

strComputer = InputBox("Enter full computer name (i.e. SWSA29565) or IP address. Leave blank to run against your own PC.")

If IsEmpty(strComputer) Then
	WScript.quit()
ElseIf	strComputer = "" Then
	strComputer = "."
End If

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = GetObject( "winmgmts://" & strComputer & "/root/cimv2" )
objSWbemServices.Security_.ImpersonationLevel = 3

' Get machine name just in case the user is using . for testing
Set colItems = objSWbemServices.ExecQuery("select * from Win32_ComputerSystem")
For Each objItem in colItems
	strMachineName = objItem.Name
Next

Set colItems = objSWbemServices.ExecQuery ("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled='True'")
For Each objItem In colItems
	strIPAddress = Join(objItem.IPAddress, ",")
Next

strFilename = strMachineName + " slow performance checks.txt"
Set objFileSystem = CreateObject("Scripting.FileSystemObject")
Set objFile = objFileSystem.OpenTextFile(strFilename, 8, True)

objFile.WriteLine("")
objFile.WriteLine("======================================================================")
objFile.WriteLine("")
objFile.WriteLine(" >> Slow Performance Checks.vbs v2.4 - Alistair McMillan")
objFile.WriteLine("[  ] Performing tests on " & strMachineName & " - " & strIPAddress)
objFile.WriteLine("[  ] Scan started at " & Now)
objFile.WriteLine("")

' Get page file sizes
Set colOSItems = objSWbemServices.ExecQuery("select * from Win32_OperatingSystem")
For Each objItem in colOSItems
	freePhysicalMemory = objItem.FreePhysicalMemory
	freeSpaceInPagingFiles = objItem.FreeSpaceInPagingFiles
	sizeStoredInPagingFiles = objItem.SizeStoredInPagingFiles
	strOperatingSystem = objItem.Caption
	strServiceParkMajor = objItem.ServicePackMajorVersion
	strServiceParkMinor = objItem.ServicePackMinorVersion
	strLastBootUpTime = WMIDateStringToDate(objItem.LastBootUpTime)
Next

' Basic info
Set colItems = objSWbemServices.ExecQuery("select * from Win32_ComputerSystem")
For Each objItem in colItems

	If (strOperatingSystem = "Microsoft Windows 2000 Professional") Then
		boolWindows2000 = True
		objFile.WriteLine("[!!] Operating System: " & strOperatingSystem)
	Else
		objFile.WriteLine("[  ] Operating System: " & strOperatingSystem)
	End If
	If ((strOperatingSystem = "Microsoft Windows XP Professional") And (strServiceParkMajor < 3)) Then
		boolServicePackMissing = True
		objFile.WriteLine("[!!] Service Pack: " & strServiceParkMajor & "." & strServiceParkMinor)
	ElseIf ((strOperatingSystem = "Microsoft Windows 2000 Professional") And (strServiceParkMajor < 4)) Then
		boolServicePackMissing = True
		objFile.WriteLine("[!!] Service Pack: " & strServiceParkMajor & "." & strServiceParkMinor)
	Else
		objFile.WriteLine("[  ] Service Pack: " & strServiceParkMajor & "." & strServiceParkMinor)
	End If		

	objFile.WriteLine("[  ] Domain: " & objItem.Domain )
	If (DateDiff("d", strLastBootUpTime, Now) > 3) Then
		boolNeedsReboot = True
		objFile.WriteLine("[!!] Last Boot Up Time: " & strLastBootUpTime )
	Else
		objFile.WriteLine("[  ] Last Boot Up Time: " & strLastBootUpTime )
	End If
	strCurrentUser = objItem.UserName
	objFile.WriteLine("[  ] User: " & objItem.UserName)
Next

objFile.WriteLine("")

' CPU
Set colItems = objSWbemServices.ExecQuery("Select * from Win32_Processor")
For Each objItem in colItems
	If ((objItem.LoadPercentage > 90) Or (Round(objItem.CurrentClockSpeed/1000, 1) < 1)) Then
		objFile.Write("[!!] ")
	Else
		objFile.Write("[  ] ")
	End If		
	objFile.WriteLine(objItem.DeviceID & " - " & _
		"Clockspeed: " & Round(objItem.CurrentClockSpeed/1000, 1) & " GHz - " & _ 
		"Load: " & objItem.LoadPercentage & "% - " & _
		"Name: " & objItem.Name)
Next

objFile.WriteLine("")

' RAM/Paging
Set colItems = objSWbemServices.ExecQuery("select * from Win32_ComputerSystem")
For Each objItem in colItems
	If (objItem.TotalPhysicalMemory/1024/1024 < 1024) Then
		objFile.WriteLine("[!!] RAM Free/Total: " & Round(freePhysicalMemory/1024, 1) & "/" & Round(objItem.TotalPhysicalMemory/1024/1024, 1) & " MB")
	Else
		objFile.WriteLine("[  ] RAM Free/Total: " & Round(freePhysicalMemory/1024, 1) & "/" & Round(objItem.TotalPhysicalMemory/1024/1024, 1) & " MB")
	End If
	objFile.WriteLine("[  ] Page File Free/Total: " & Round(freeSpaceInPagingFiles/1024, 1) & "/" & Round(sizeStoredInPagingFiles/1024, 1) & " MB")
Next
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management"
strValueName = "PagingFiles"
objRegistry.GetMultiStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, values
If (IsNull(values)) Then
	objFile.WriteLine("[!!] No paging files key")
Else
	For Each strValue in values
		If (InStr(strValue, " 0 0")) Then
			objFile.WriteLine("[  ] Virtual memory is SYSTEM MANAGED - " & strValue)
		Else
			boolVirtMemNotSystemManaged = True
			objFile.WriteLine("[!!] Virtual memory is NOT system managed - " & strValue)
		End If
	Next
End If

objFile.WriteLine("")

' Lists disks with total/free space
Set colItems = objSWbemServices.ExecQuery("select * from Win32_LogicalDisk")
For Each objItem in colItems
	If ("C:" = objItem.Name) Then
		If (IsNull(objItem.Size) Or IsNull(objItem.FreeSpace)) Then
			If (objItem.FreeSpace/objItem.Size < 10) Then
				objFile.Write("[!!] ")
			Else
				objFile.Write("[  ] ")
			End If
			objFile.WriteLine(objItem.Name & " " & objItem.FreeSpace/objItem.Size & " % FREE - Free/Total Space: "  & objItem.FreeSpace/1024/1024/1024 & " GB / " & objItem.Size/1024/1024/1024& " GB")
		else
			If (Round((objItem.FreeSpace/objItem.Size)*100, 1) < 10) Then
				objFile.Write("[!!] ")
			Else
				objFile.Write("[  ] ")
			End If
			objFile.WriteLine(objItem.Name & " " & Round((objItem.FreeSpace/objItem.Size)*100, 1) & " % FREE - " & " Free/Total Space: " & Round(objItem.FreeSpace/1024/1024/1024, 2) & " GB / " & Round(objItem.Size/1024/1024/1024, 2) & " GB ")
		End If
	End If
Next

objFile.WriteLine("")

objFile.WriteLine("RUNNING PROCESSES")
objFile.WriteLine("-----------------")

' RAM
objFile.WriteLine(vbTab & " WORKING   PAGE FILE   PROCESS")
objFile.WriteLine(vbTab & " SET (MB), USAGE (MB), NAME         , COMMAND")

' Creating disconnected recordset to hold data
Const adDouble = 5
Const adVarChar = 200
Const MaxCharacters = 511
Set DataList = CreateObject("ADOR.Recordset")
DataList.Fields.Append "WorkingSetSize", adDouble, MaxCharacters
DataList.Fields.Append "PageFileUsage", adDouble, MaxCharacters
DataList.Fields.Append "Name", adVarChar, MaxCharacters
DataList.Fields.Append "CommandLine", adVarChar, MaxCharacters
DataList.Open

' Retrieve list of processes and load into recordset
Set colItems = objSWbemServices.ExecQuery("select * from Win32_Process")
For Each objItem in colItems
	DataList.AddNew
	DataList("WorkingSetSize") = CDbl(objItem.WorkingSetSize)
	DataList("PageFileUsage") = CDbl(objItem.PageFileUsage)
	DataList("Name") = objItem.Name
	' Because Windows 2000 doesn't return command line for processes
	If (boolWindows2000) Then
	Else
		' Because command line can sometimes be null and recordsets don't like null values
		If (IsNull(objItem.CommandLine)) Then
			DataList("CommandLine") = "-"
		Else
			DataList("CommandLine") = objItem.CommandLine
		End If
	End If
	DataList.Update
Next
DataList.Sort = "PageFileUsage DESC"
DataList.MoveFirst
Do Until DataList.EOF
	objFile.Write("[  ]  ")

	objFile.Write(PadNumbers(Round(DataList.Fields.Item("WorkingSetSize")/1024/1024, 2)))
	objFile.Write(FormatNumber(Round(DataList.Fields.Item("WorkingSetSize")/1024/1024, 2), 2, -1))
	objFile.Write(",    ")

	objFile.Write(PadNumbers(Round(DataList.Fields.Item("PageFileUsage")/1024/1024, 2)))
	objFile.Write(FormatNumber(Round(DataList.Fields.Item("PageFileUsage")/1024/1024, 2), 2, -1) & ", " & DataList.Fields.Item("Name"))
	If (boolWindows2000) Then
	Else
		objFile.Write(", " & DataList.Fields.Item("CommandLine") )
	End If
	objFile.WriteLine("")
	DataList.MoveNext
Loop

objFile.WriteLine("")

'	objFile.WriteLine("MAPPED DRIVES")
'	objFile.WriteLine("-------------")
'	
'	Set colItems = objSWbemServices.ExecQuery("Select * from Win32_MappedLogicalDisk")
'
'	For Each objItem in colItems
'		strOutput = strOutput & "Name: " & objItem.Name & vbCr
'		strOutput = strOutput & "Provider Name: " & objItem.ProviderName & vbCr
'		Wscript.Echo strOutput
'		strOutput = ""
'	Next
'	
'	objFile.WriteLine("")

objFile.WriteLine("SCHEDULED TASKS")
objFile.WriteLine("---------------")
objFile.WriteLine("")

SchTasksCommand = "schtasks /query /V /FO TABLE /s " & strMachineName
Set oExec = WshShell.Exec(SchTasksCommand)

'	On Error Resume Next
Do While oExec.StdOut.AtEndOfStream <> True
	Line = oExec.StdOut.ReadLine
	If Not IsNull(Line) And (Len(Line) > 0) Then
		objFile.WriteLine("[  ] " & Line)
	End If
Loop

objFile.WriteLine("")

objFile.WriteLine("PATH")
objFile.WriteLine("----")
objFile.WriteLine("")

Set colItems = objSWbemServices.ExecQuery("Select * from Win32_Environment")

For Each objItem in colItems
	If (objItem.Name = "PATH" or objItem.Name = "Path") Then
		objFile.WriteLine("[  ] Name: " & objItem.Name)
		If (InStr(objItem.VariableValue,"\\")) Then
			boolRemotePathProblem = True
			objFile.Write("[!!] ")
		Else
			objFile.Write("[  ] ")
		End If
		objFile.WriteLine("Variable Value: " & objItem.VariableValue)
	End If
Next

objFile.WriteLine("")

objFile.WriteLine("PRINTERS")
objFile.WriteLine("--------")
objFile.WriteLine("")

Set colPrinters = objSWbemServices.ExecQuery("Select * From Win32_Printer")

For Each objPrinter in colPrinters
	If (InStr(objPrinter.Name, "CORP-FP1") Or InStr(objPrinter.Name, "corp-fp1") Or InStr(objPrinter.Name, "Corp-fp1")) Then
		boolAncientPrinterQueues = True
		objFile.Write("[!!] ")	
	Else
		objFile.Write("[  ] ")
	End If
	If objPrinter.Attributes And 64 Then 
		strPrinterType = "Local   -- "
	Else
		strPrinterType = "Network -- "
	End If
	objFile.WriteLine strPrinterType & objPrinter.Name 
Next

objFile.WriteLine("")

objFile.WriteLine("MACHINE STARTUP ITEMS")
objFile.WriteLine("---------------------")

Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
objRegistry.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes
For index=0 To UBound(arrValueNames)
	Select Case arrValueTypes(index)
		Case REG_SZ
			objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames(index), strValue
			if(InStr(1, strValue, "\\") = 1) Then
				boolRemoteStartupItem = True
				objFile.Write("[!!] ")
			Else
				objFile.Write("[  ] ")
			End If
			objFile.WriteLine arrValueNames(index) & ",  " & strValue
		Case REG_EXPAND_SZ
			objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames(index), strValue
			if(InStr(1, strValue, "\\") = 1) Then
				boolRemoteStartupItem = True
				objFile.Write("[!!] ")
			Else
				objFile.Write("[  ] ")
			End If
			objFile.WriteLine arrValueNames(index) & ",  " & strValue
		Case REG_BINARY
			' Should never reach here
		Case REG_DWORD
			' Should never reach here
		Case REG_MULTI_SZ
			' Should never reach here
	End Select 
Next

objFile.WriteLine("")

objFile.WriteLine("USER STARTUP ITEMS")
objFile.WriteLine("------------------")

Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
If (IsNull(strCurrentUser)) Then
	objFile.WriteLine "No user currently logged in"
Else
	objRegistry.EnumValues HKEY_CURRENT_USER, strKeyPath, arrValueNames, arrValueTypes
	On Error Resume Next
	For index=0 To UBound(arrValueNames)
		Select Case arrValueTypes(index)
			Case REG_SZ
				objRegistry.GetStringValue HKEY_CURRENT_USER, strKeyPath, arrValueNames(index), strValue
				if(InStr(1, strValue, "\\") = 1) Then
					boolRemoteStartupItem = True
					objFile.Write("[!!] ")
				Else
					objFile.Write("[  ] ")
				End If
				objFile.WriteLine arrValueNames(index) & ",  " & strValue
			Case REG_EXPAND_SZ
				objRegistry.GetStringValue HKEY_CURRENT_USER, strKeyPath, arrValueNames(index), strValue
				if(InStr(1, strValue, "\\") = 1) Then
					boolRemoteStartupItem = True
					objFile.Write("[!!] ")
				Else
					objFile.Write("[  ] ")
				End If
				objFile.WriteLine arrValueNames(index) & ",  " & strValue
			Case REG_BINARY
				' Should never reach here
			Case REG_DWORD
				' Should never reach here
			Case REG_MULTI_SZ
				' Should never reach here
		End Select 
	Next
	On Error Goto 0
End If

objFile.WriteLine("")

objFile.WriteLine("TEMP FOLDERS")
objFile.WriteLine("------------")
objFile.WriteLine("FOLDER SIZE (MB), FOLDER NAME")

If objFileSystem.FolderExists("\\" & strMachineName & "\c$\windows\temp") Then
	Set queriesFolder = objFileSystem.GetFolder("\\" & strMachineName & "\c$\windows\temp")
	On Error Resume Next
	tempFolderTotal = tempFolderTotal + Round(queriesFolder.Size/1024/1024, 1)
	objFile.WriteLine("[  ] " & PadNumbers(Round(queriesFolder.Size/1024/1024, 1)) & FormatNumber(Round(queriesFolder.Size/1024/1024, 1), 2, -1) & "MB, <" & queriesFolder.Path & ">")
	On Error Goto 0
End If

If objFileSystem.FolderExists("\\" & strMachineName & "\c$\temp") Then
	Set queriesFolder = objFileSystem.GetFolder("\\" & strMachineName & "\c$\temp")
	On Error Resume Next
	tempFolderTotal = tempFolderTotal + Round(queriesFolder.Size/1024/1024, 1)
	objFile.WriteLine("[  ] " & PadNumbers(Round(queriesFolder.Size/1024/1024, 1)) & FormatNumber(Round(queriesFolder.Size/1024/1024, 1), 2, -1) & "MB, <" & queriesFolder.Path & ">")
	On Error Goto 0
End If

If objFileSystem.FolderExists("\\" & strMachineName & "\c$\windows\system32\config\systemprofile\local settings\temp") Then
	Set queriesFolder = objFileSystem.GetFolder("\\" & strMachineName & "\c$\windows\system32\config\systemprofile\local settings\temp")
	On Error Resume Next
	tempFolderTotal = tempFolderTotal + Round(queriesFolder.Size/1024/1024, 1)
	objFile.WriteLine("[  ] " & PadNumbers(Round(queriesFolder.Size/1024/1024, 1)) & FormatNumber(Round(queriesFolder.Size/1024/1024, 1), 2, -1) & "MB, <" & queriesFolder.Path & ">")
	On Error Goto 0
End If

Set objFolder = objFileSystem.GetFolder("\\" & strMachineName & "\c$\Documents and Settings")
Set objSubFolders = objFolder.SubFolders
For Each objSubFolder in objSubFolders
	If objFileSystem.FolderExists("\\" & strMachineName & "\c$\documents and settings\" & objSubFolder.name & "\local settings\temp") Then
		Set queriesFolder = objFileSystem.GetFolder("\\" & strMachineName & "\c$\documents and settings\" & objSubFolder.name & "\local settings\temp")
		On Error Resume Next
		If (InStr(objSubFolder.name, "Default User")) And (Round(queriesFolder.Size/1024/1024, 1) > 10) Then
			boolDefaultTempProblem = True
		End If
		tempFolderTotal = tempFolderTotal + Round(queriesFolder.Size/1024/1024, 1)
		objFile.WriteLine("[  ] " & PadNumbers(Round(queriesFolder.Size/1024/1024, 1)) & FormatNumber(Round(queriesFolder.Size/1024/1024, 1), 2, -1) & "MB, <" & queriesFolder.Path & ">")
		On Error Goto 0
	End If
Next

If (tempFolderTotal > 1024) Then
	boolTempFoldersProblem = True
	objFile.Write("[!!] ")
Else
	objFile.Write("[  ] ")
End If
objFile.WriteLine(PadNumbers(tempFolderTotal) & FormatNumber(tempFolderTotal, 2, -1) & "MB TOTAL")

objFile.WriteLine("")

objFile.WriteLine("TEMPORARY INTERNET FOLDERS")
objFile.WriteLine("--------------------------")
objFile.WriteLine("FOLDER SIZE (MB), FOLDER NAME")

If objFileSystem.FolderExists("\\" & strMachineName & "\c$\windows\system32\config\systemprofile\local settings\temporary internet files") Then
	Set queriesFolder = objFileSystem.GetFolder("\\" & strMachineName & "\c$\windows\system32\config\systemprofile\local settings\temporary internet files")
	On Error Resume Next
	temporaryInternetFolderTotal = temporaryInternetFolderTotal + Round(queriesFolder.Size/1024/1024, 1)
	objFile.WriteLine("[  ] " & PadNumbers(Round(queriesFolder.Size/1024/1024, 1)) & FormatNumber(Round(queriesFolder.Size/1024/1024, 1), 2, -1) & "MB, <" & queriesFolder.Path & ">")
	On Error Goto 0
End If

Set objFolder = objFileSystem.GetFolder("\\" & strMachineName & "\C$\documents and settings")
Set objSubFolders = objFolder.SubFolders
For Each objSubFolder in objSubFolders
	If objFileSystem.FolderExists("\\" & strMachineName & "\c$\documents and settings\" & objSubFolder.name & "\local settings\temporary internet files") Then
		Set queriesFolder = objFileSystem.GetFolder("\\" & strMachineName & "\c$\documents and settings\" & objSubFolder.name & "\local settings\temporary internet files")
		On Error Resume Next
		If (InStr(objSubFolder.name, "Default User")) And (Round(queriesFolder.Size/1024/1024, 1) > 10) Then
			boolDefaultTempInternetFilesProblem = True
		End If
		temporaryInternetFolderTotal = temporaryInternetFolderTotal + Round(queriesFolder.Size/1024/1024, 1)
		objFile.WriteLine("[  ] " & PadNumbers(Round(queriesFolder.Size/1024/1024, 1)) & FormatNumber(Round(queriesFolder.Size/1024/1024, 1), 2, -1) & "MB, <" & queriesFolder.Path & ">")
		On Error Goto 0
	End If
Next

If (temporaryInternetFolderTotal > 1024) Then
	boolTempInternetFoldersProblem = True
	objFile.Write("[!!] ")
Else
	objFile.Write("[  ] ")
End If
objFile.WriteLine(PadNumbers(temporaryInternetFolderTotal) & FormatNumber(temporaryInternetFolderTotal, 2, -1) & "MB TOTAL")

objFile.WriteLine("")

objFile.WriteLine("QUERIES FOLDERS")
objFile.WriteLine("---------------")
objFile.WriteLine("FOLDER SIZE (MB), FOLDER NAME")

Set objFolder = objFileSystem.GetFolder("\\" & strMachineName & "\C$\documents and settings")
Set objSubFolders = objFolder.SubFolders
For Each objSubFolder in objSubFolders
	If objFileSystem.FolderExists("\\" & strMachineName & "\c$\documents and settings\" & objSubFolder.name & "\application data\microsoft\queries") Then
		Set queriesFolder = objFileSystem.GetFolder("\\" & strMachineName & "\c$\documents and settings\" & objSubFolder.name & "\application data\microsoft\queries")
		On Error Resume Next
		queriesFolderTotal = queriesFolderTotal + Round(queriesFolder.Size/1024/1024, 1)
		If (Round(queriesFolder.Size/1024/1024, 1) > 1) Then
			hasQueriesProblem = True
		End If
		objFile.WriteLine("[  ] " & PadNumbers(Round(queriesFolder.Size/1024/1024, 1)) & FormatNumber(Round(queriesFolder.Size/1024/1024, 1), 2, -1) & "MB, <" & queriesFolder.Path & ">")
		On Error Goto 0
	End If
Next

If (hasQueriesProblem) Then
	objFile.Write("[!!] ")
Else
	objFile.Write("[  ] ")
End If
objFile.WriteLine(PadNumbers(queriesFolderTotal) & FormatNumber(queriesFolderTotal, 2, -1) & "MB TOTAL")

objFile.WriteLine("")

objFile.WriteLine("RECOMMENDATIONS")
objFile.WriteLine("---------------")
If (boolWindows2000) Then
	objFile.WriteLine(vbTab & " Machine is running Windows 2000 - unsupported by Microsoft")
End If
If (boolServicePackMissing) Then
	objFile.WriteLine(vbTab & " Install missing Service Pack")
End If
If (boolVirtMemNotSystemManaged) Then
	objFile.WriteLine(vbTab & " Change Virtual Memory to System Managed")
End If
If (boolNeedsReboot) Then
	objFile.WriteLine(vbTab & " Reboot machine - hasn't been rebooted for " & DateDiff("d", strLastBootUpTime, Now) & " days.")
End If
If (boolRemoteStartupItem) Then
	objFile.WriteLine(vbTab & " Remove Startup item that launches from remote server")
End If
if (boolRemotePathProblem) Then
	objFile.WriteLine(vbTab & " System path has a remote path, this may cause slowdown")
End If
If (boolAncientPrinterQueues) Then
	objFile.WriteLine(vbTab & " Remove any printers queues that are hosted on CORP-FP1")
End If
If (boolTempFoldersProblem) Then
	objFile.WriteLine(vbTab & " Clear Temp folders")
End If
If (boolDefaultTempProblem) Then
	objFile.WriteLine(vbTab & " Clear Default User profile's Temp folder")
End If
If (boolTempInternetFoldersProblem) Then
	objFile.WriteLine(vbTab & " Clear Temporary Internet Folders")
End If
If (boolDefaultTempInternetFilesProblem) Then
	objFile.WriteLine(vbTab & " Clear Default User profile's Temporary Internet Files folder")
End If
If (hasQueriesProblem) Then
	objFile.WriteLine(vbTab & " Clear Queries folders and Excel fix may need to be applied")
End If

objFile.WriteLine("")
objFile.WriteLine("[  ] Scan finished at " & Now)

SchTasksCommand = "notepad.exe " & strFilename
WshShell.Exec(SchTasksCommand)

Wscript.Quit
