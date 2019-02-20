'vbscript
dim file1
dim version
dim verNum
dim WshShell
Set WshShell = CreateObject("WScript.Shell")

Set shell = CreateObject("WScript.Shell")
Set getOSVersion = shell.exec("%comspec% /c ver")
version = getOSVersion.stdout.readall
wscript.echo version
Select Case True
   Case InStr(version, "n 5.") > 1 : GetOS = "XP"
   Case InStr(version, "n 6.") > 1 : GetOS = "Vista"
   Case Else : GetOS = "Unknown"
End Select
wscript.echo GetOS
