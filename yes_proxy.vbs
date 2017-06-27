dim oShell
set oShell = Wscript.CreateObject("Wscript.Shell")
oShell.RegWrite "HKCU\Software\Microsoft\Windows\currentVersion\Internet Settings\ProxyServer", "192.168.0.1:3128", "REG_SZ"
Set oShell = Nothing