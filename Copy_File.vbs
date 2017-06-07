
Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

objFSO.CopyFile "C:\test.txt", "C:\Windows\Desktop\A"

Set objFSO = Nothing