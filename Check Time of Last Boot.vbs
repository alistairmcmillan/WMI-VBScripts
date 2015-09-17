' Alistair McMillan
' Start Date: 10 August 2011
' -----------------------------------------------

Option Explicit

' Declare variables
Dim colItems, objItem, objWMIService
Dim strBoot, strBootDate, strBootDay, strBootHour, strBootMins
Dim strBootMonth, strBootTime, strBootYear, strComputer, strMsg, strQuery

strComputer = InputBox("Enter full computer name (i.e. SWSA29565) or IP address. Leave blank to run against your own PC.")

If IsEmpty(strComputer) Then
	WScript.quit()
ElseIf	strComputer = "" Then
	strComputer = "."
End If

' Connect to specified computer
Set objWMIService = GetObject( "winmgmts://" & strComputer & "/root/cimv2" )
Set colItems = objWMIService.ExecQuery( "Select * from Win32_OperatingSystem", , 48 )
For Each objItem in colItems
	strBootYear  = Left( objItem.LastBootUpTime, 4 )
	strBootMonth = Mid( objItem.LastBootUpTime,  5, 2 )
	strBootDay   = Mid( objItem.LastBootUpTime,  7, 2 )
	strBootDate  = DateValue( strBootDay & "-" & strBootMonth & "-" & strBootYear )
	strBootHour  = Mid( objItem.LastBootUpTime,  9, 2 )
	strBootMins  = Mid( objItem.LastBootUpTime, 11, 2 )
	strBootTime  = strBootHour & ":" & strBootMins
	strBoot = strBootDate & ", " & strBootTime
	strMsg  = "Last boot time of " & strComputer & ": " & strBoot
Next

' Display results
WScript.Echo strMsg

'Done
WScript.Quit(0)
