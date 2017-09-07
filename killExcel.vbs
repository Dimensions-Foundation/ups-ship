Dim objWMIService, WshShell
Dim proc, procList, killCount
Dim strComputer, strCommand
killCount = 0
strCommand = "taskkill /F /IM excel.exe"
strComputer = "."
Set WshShell = WScript.CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:"& "{impersonationLevel=impersonate}!\\"& strComputer & "\root\cimv2")
Set procList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'excel.exe'")
For Each proc In procList
	WshShell.run strCommand, 0, TRUE
	killCount = killCount + 1
Next
MsgBox killCount & " Excel Processes Terminated"
