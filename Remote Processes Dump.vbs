' Remote Processes Dump
' Author: Alistair McMillan
' Start Date: 31 October 2012
' -----------------------------------------------' 
Option Explicit
Dim strComputer, strFilename, objFileSystem, objFile, objSWbemLocator, objSWbemServices, colItems, objItem, strMachineName

strComputer = InputBox("Enter full computer name (i.e. SWSA40115) or IP address")

Set objFileSystem = CreateObject("Scripting.FileSystemObject")

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = GetObject( "winmgmts://" & strComputer & "/root/cimv2" )
objSWbemServices.Security_.ImpersonationLevel = 3

' Basic info
Set colItems = objSWbemServices.ExecQuery("select * from Win32_ComputerSystem")
For Each objItem in colItems
	strMachineName = objItem.Name
Next

strFilename = strMachineName + ".csv"
Set objFile = objFileSystem.OpenTextFile(strFilename, 8, True)

' RAM
objFile.WriteLine("MACHINE, PROCESS ID, PROCESS NAME, WORKING SET (MB), PAGE FILE USAGE (MB), COMMAND")

Set colItems = objSWbemServices.ExecQuery("select * from Win32_Process")
For Each objItem in colItems
	objFile.WriteLine(strMachineName & ", " & objItem.ProcessId  & ", " & objItem.Name& ", " & Round(objItem.WorkingSetSize/1024/1024, 2) & ", " & Round(objItem.PageFileUsage/1024/1024, 2) & ", " & objItem.CommandLine )
Next

objFile.Close

set objFile = NOTHING
set objFileSystem = NOTHING

Wscript.Echo "DONE. Created " & strFilename

WSCript.Quit
