strComputer = "."

Set objGroup = GetObject("WinNT://"&strComputer&"/Administrators")

For Each objUser In objGroup.Members
    If objUser.Name <> "Administrator" AND objUser.Name <> "Domain Admins" Then
        objGroup.Remove(objUser.AdsPath)
    End If
Next

Wscript.Echo "OK!"