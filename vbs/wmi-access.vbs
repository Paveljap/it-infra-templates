Dim objWMIService

Sub ListHardware()
	Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

	For Each objItem in colItems
		Wscript.Echo "NumbersOfProcessors" & objItem.NumberOfProcessors
		Wscript.Echo "TotalPhysicalMemory"& objItem.TotalPhysicalMemory
		Wscript.Echo "Name" & objItem.Name
		

	Next
End Sub
'-------------------------------------------------------------------------------
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	
' Оборудование	
ListHardware