Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("Shell.Application")

Path = oShell.BrowseForFolder(0,"Choose Folder",0,17).Items.Item.Path
Name = InputBox("Please Enter A FileName",,"FileName.txt")

If Path <> "" And Name <> "" Then
	Set TextFile = oFSO.OpenTextFile(Path & Name,2,True)
End If

'Write something in the file.
TextFile.WriteLine "My Save As..."
TextFile.Close 