Set shell = CreateObject("WScript.Shell")
scriptPath = Replace(WScript.ScriptFullName, "SourceCodeLauncher.vbs", "sourceCode.ps1")
command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """"
shell.Run command, 0, False
