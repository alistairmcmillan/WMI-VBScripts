' Alistair McMillan
' Start Date: 24 June 2014
' -----------------------------------------------

Option Explicit

Dim strComputer, strFilename, objFileSystem, objFile, objFolder, objShell, objFolderItem, x

strComputer = InputBox("Enter full computer name (i.e. SWSA29565) or IP address. Leave blank to run against your own PC.")

If IsEmpty(strComputer) Then
	WScript.quit()
ElseIf	strComputer = "" Then
	strComputer = "."
End If

Set objFileSystem = CreateObject("Scripting.FileSystemObject")
If objFileSystem.FolderExists("\\" & strComputer & "\c$\Program Files\SAP\FrontEnd\SAPgui" ) Then
	If objFileSystem.FileExists("\\" & strComputer & "\c$\Program Files\SAP\FrontEnd\SAPgui\SAPgui.exe") Then
		Dim arrDetails(100)
		Set objShell = CreateObject("Shell.Application")
		Set objFolder = objShell.Namespace("\\" & strComputer & "\c$\Program Files\SAP\FrontEnd\SAPgui\")
		Set objFolderItem = objFolder.ParseName("SAPgui.exe")
		For x = 0 To 100
			arrDetails(x) = objFolder.GetDetailsOf(objFolder.Items, x)
			If lcase(arrDetails(x))= "product version" Then
				MsgBox(strComputer & " has SAP GUI version " & objFolder.GetDetailsOf(objFolderItem, x))
				Exit For
			End If
		Next
	Else
		MsgBox("Uh oh! " & strComputer & " doesn't seem to have SAPgui.exe installed.")
	End If
Else
	MsgBox("Uh oh! " & strComputer & " doesn't even seem to have a SAPgui folder.")
End If

Wscript.Quit
