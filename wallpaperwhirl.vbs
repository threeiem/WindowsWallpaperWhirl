dim wsShell
dim objShell

On Error Resume Next

strComputer = "."
set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\wmi")
set wmiItems = objWMIService.ExecQuery("SELECT * FROM WMIMonitorID")
set objShell = CreateObject("shell.application")
set wsShell = WScript.CreateObject("WScript.Shell")

objShell.MinimizeAll

For Each objItem In wmiItems

    wsShell.AppActivate "Program Manager"
    wsShell.SendKeys("{F5}")   
    wsShell.SendKeys "^ "
    wsShell.SendKeys "+{F10}"
    wsShell.SendKeys "n"

    WScript.Sleep(1111)
    
Next

objShell.UndoMinimizeAll

set objWMIService = Nothing
set colItems = Nothing
set objShell = Nothing
set wsShell = Nothing

