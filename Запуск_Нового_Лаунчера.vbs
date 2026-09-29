Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strDir = fso.GetParentFolderName(WScript.ScriptFullName)
strCmd = fso.BuildPath(strDir, fso.GetBaseName(WScript.ScriptFullName) & ".cmd")
WshShell.CurrentDirectory = strDir
WshShell.Environment("PROCESS")("FENG_SILENT") = "1"
WshShell.Run "cmd /c """"" & strCmd & """""", 0, False
