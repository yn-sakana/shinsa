Set shell = CreateObject("WScript.Shell")
shell.Run "powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & Replace(WScript.ScriptFullName, "shinsa.vbs", "gui\Start-Gui.ps1") & """", 0, False
