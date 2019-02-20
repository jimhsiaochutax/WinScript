'vbscript
dim file1
dim version
dim verNum
dim WshShell
Set WshShell = CreateObject("WScript.Shell")

FILE1 = "C:\Program Files\LibreOffice 5\program\swriter.exe"
version = ("5.1")
Set objFSO = CreateObject("Scripting.FileSystemObject")

verNum = objFSO.GetFileVersion(FILE1)
If verNum < version Then
'If Left(verNum,5)<3.0.5 then <path_to_executable>
'WshShell.run(path to executable)
Wscript.Echo verNum
End If

