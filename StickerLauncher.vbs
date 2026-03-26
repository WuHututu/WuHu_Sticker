Set shell = CreateObject("WScript.Shell")
scriptPath = Replace(WScript.ScriptFullName, "StickerLauncher.vbs", "Sticker.ps1")
command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """"
shell.Run command, 0, False
