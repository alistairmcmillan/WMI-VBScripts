' Alistair McMillan
' Start Date: 31 August 2011
' -----------------------------------------------

Option Explicit

' Declare variables
Dim intValidArgs
Dim colItems, objItem, objWMIService
Dim strBoot, strBootDate, strBootDay, strBootHour, strBootMins
Dim strBootMonth, strBootTime, strBootYear, strComputer, strMsg, strQuery

intValidArgs = 0

strComputer = InputBox("Enter full computer name (i.e. SWSA29565) or IP address. Leave blank to run against your own PC.")

If IsEmpty(strComputer) Then
	WScript.quit()
ElseIf	strComputer = "" Then
	strComputer = "."
End If

' Connect to local computer
Set objWMIService = GetObject( "winmgmts://./root/cimv2" )
Set colItems = objWMIService.ExecQuery( "Select * from Win32_LocalTime" )
For Each objItem in colItems
	strMsg =  "Time on local PC: " _
				& objItem.Hour & ":" _
				& objItem.Minute & ":" _ 
				& objItem.Second & " " _
				& objItem.Day & "/" _
				& objItem.Month & "/" _
				& objItem.Year
Next

' Connect to specified computer
Set objWMIService = GetObject( "winmgmts://" & strComputer & "/root/cimv2" )
Set colItems = objWMIService.ExecQuery( "Select * from Win32_LocalTime" )
For Each objItem in colItems
	strMsg =  strMsg & VbCr & VbCr _
				& "Time on " & strComputer &  ": " _
				& objItem.Hour & ":" _
				& objItem.Minute & ":" _ 
				& objItem.Second & " " _
				& objItem.Day & "/" _
				& objItem.Month & "/" _
				& objItem.Year
Next

' Display results
WScript.Echo strMsg

WScript.Quit(0)
