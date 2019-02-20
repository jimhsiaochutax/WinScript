IEKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\"
Set WshShell = CreateObject("WScript.Shell")
On Error Resume Next
svcVersion = Null
svcVersion = WshShell.RegRead(IEKeyPath & "svcVersion")
version = Null
version = WshShell.RegRead(IEKeyPath & "Version")
On Error GOTO 0
Set WshShell = Nothing

WScript.echo version & " ; " & svcVersion
If left(version,1) < 9 Then
msgbox version & " ; " & svcVersion
End If

