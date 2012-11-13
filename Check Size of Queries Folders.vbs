' Alistair McMillan
' Start Date: 06 November 2012
' -----------------------------------------------

Option Explicit

Dim strComputer, strFilename, objFileSystem, objFile, objFolder, _
	objSubFolders, objSubFolder, queriesFolder

strComputer = InputBox("Enter full computer name (i.e. SWSA29565) or IP address")

strFilename = strComputer + " queries folders.csv"
Set objFileSystem = CreateObject("Scripting.FileSystemObject")
Set objFile = objFileSystem.OpenTextFile(strFilename, 8, True)

Set objFolder = objFileSystem.GetFolder("\\" & strComputer & "\C$\documents and settings")
Set objSubFolders = objFolder.SubFolders

objFile.WriteLine("FOLDER SIZE (MB), FOLDER NAME")

For Each objSubFolder in objSubFolders
	If objFileSystem.FolderExists("\\" & strComputer & "\c$\documents and settings\" & objSubFolder.name & "\application data\microsoft\queries") Then
		Set queriesFolder = objFileSystem.GetFolder("\\" & strComputer & "\c$\documents and settings\" & objSubFolder.name & "\application data\microsoft\queries")
		On Error Resume Next
		objFile.Write(Round(queriesFolder.Size/1024/1024, 1))
		objFile.Write(", " & queriesFolder.Path)
		objFile.WriteLine(" ")
	End If
Next

Wscript.Echo "Finished writing to " & strFilename
objFile.Close

set objFile = NOTHING
set objFileSystem = NOTHING

Wscript.Quit
