dim oShell
set oShell = Wscript.CreateObject("Wscript.Shell")
oShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD"
Set oShell = Nothing
