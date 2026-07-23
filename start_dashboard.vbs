Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
projectDir = fso.GetParentFolderName(WScript.ScriptFullName)
pythonExe = "C:\Users\admin\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
command = "cmd.exe /c cd /d """ & projectDir & """ && """ & pythonExe & """ server.py >> server.out.log 2>> server.err.log"
shell.Run command, 0, False
