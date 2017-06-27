dim oShell
set oShell = Wscript.CreateObject("Wscript.Shell")

if msgbox("Habilitar Proxy?", vbQuestion or vbYesNo) = vbYes then
oShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1, "REG_DWORD"
oShell.RegWrite "HKCU\Software\Microsoft\Windows\currentVersion\Internet Settings\ProxyServer", "192.168.0.1:3128", "REG_SZ"
else
oShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD"

End if

Set oShell = Nothing