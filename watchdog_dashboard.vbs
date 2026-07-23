Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

projectDir = fso.GetParentFolderName(WScript.ScriptFullName)
pythonExe = "C:\Users\admin\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
serverUrl = "http://127.0.0.1:8765/api/events"

Sub StartServer()
  command = "cmd.exe /c cd /d """ & projectDir & """ && """ & pythonExe & """ server.py >> server.out.log 2>> server.err.log"
  shell.Run command, 0, False
End Sub

Function IsServerReady()
  On Error Resume Next
  Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  http.setTimeouts 1000, 1000, 2000, 2000
  http.open "GET", serverUrl, False
  http.send
  IsServerReady = (Err.Number = 0 And http.Status = 200)
  Err.Clear
  On Error GoTo 0
End Function

Do
  If Not IsServerReady() Then
    StartServer
    WScript.Sleep 8000
  End If
  WScript.Sleep 30000
Loop
