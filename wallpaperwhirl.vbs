dim wmiService
dim wmiMonitorResult
dim wsShell
dim objShell

strComputer = "."
set wmiService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\wmi")
set wmiMonitorResult = wmiService.ExecQuery("SELECT * FROM WMIMonitorID")
set objShell = CreateObject("shell.application")
set wsShell = WScript.CreateObject("WScript.Shell")

' objShell.MinimizeAll

For Each monitor In wmiMonitorResult

    ' Give Desktop focus
    wsShell.AppActivate "Program Manager"
    wsShell.SendKeys("{F5}")   

    ' Use alternate context menu for the Desktop to select the 'Next Desktop Background'
    wsShell.SendKeys "^ "
    wsShell.SendKeys "+{F10}"
    wsShell.SendKeys "n"

    ' TODO: Test this
    WScript.Sleep(35)
    
Next

' objShell.UndoMinimizeAll

set objWMIService = Nothing
set colItems = Nothing
set objShell = Nothing
set wsShell = Nothing

