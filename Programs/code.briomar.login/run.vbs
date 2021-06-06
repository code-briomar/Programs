Set WshShell = CreateObject("WScript.Shell") 
WshShell.Run chr(34) & "run_this_jar.cmd" & Chr(34), 0
Set WshShell = Nothing