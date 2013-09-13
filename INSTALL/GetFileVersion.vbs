Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
WScript.Echo objFSO.GetFileVersion (WScript.Arguments(0))
WScript.Quit (0)