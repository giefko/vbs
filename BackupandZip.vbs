Set objFS = CreateObject("Scripting.FileSystemObject")
'set the .shared file source folder
strFolder = "C:\ProgramData\QlikTech\Documents"

strBackupFolder = "C:\ProgramData\QlikTech\SharedBackups\"

'Set objFolder = objFS.GetFolder(strFolder)
current=Now
mth = Month(current)
d = Day(current)
yr=Year(current)
If Len(mth) <2 Then
    mth="0"&mth
End If
If Len(d) < 2 Then
    d = "0"&d
End If
timestamp=yr & "-" & mth &"-"& d

'check the date backup folder exists, if not create it
if not objFS.FolderExists("C:\SharedBackups\" ) then
    objFS.CreateFolder("C:\SharedBackups\" & timestamp)
end if

'copy all the .shared files to the backup date folder
objFS.CopyFile strFolder & "\*.shared", "C:\SharedBackups\" , true
'


'zip the files here, using 7-zip - if you don't want to zip the files, delete the next three lines.

Dim objShell: Set objShell = CreateObject("WScript.Shell")

Command = """C:\Program Files\7-zip\7z.exe"" a " & "C:\SharedBackups\" & timestamp & ".zip " & "C:\SharedBackups\" & timestamp & "\*.shared"
RetVal = objShell.Run(Command,0,true)

'delete the backup folder - MAKE SURE THIS NEVER POINTS AT THE SOURCE OR LIVE FOLDERS!
objFS.DeleteFolder("C:\SharedBackups\" & timestamp)

'*** End Zip ***

set objFS = nothing

'*** VBScript End ***