'The sample scripts are not supported under any Microsoft standard support 
'program or service. The sample scripts are provided AS IS without warranty  
'of any kind. Microsoft further disclaims all implied warranties including,  
'without limitation, any implied warranties of merchantability or of fitness for 
'a particular purpose. The entire risk arising out of the use or performance of  
'the sample scripts and documentation remains with you. In no event shall 
'Microsoft, its authors, or anyone else involved in the creation, production, or 
'delivery of the scripts be liable for any damages whatsoever (including, 
'without limitation, damages for loss of business profits, business interruption, 
'loss of business information, or other pecuniary loss) arising out of the use 
'of or inability to use the sample scripts or documentation, even if Microsoft 
'has been advised of the possibility of such damages.

' elevate privilege to execute the code
If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
  WScript.Quit
End If

Function Normalize_VersionNumber(version)

	Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Global = true
	objRegEx.Pattern = "^(5.00)|(6.00)|(7.00)|(8.00)"

	If objRegEx.Test(version) Then
		objRegEx.Pattern = "^(\d).00"
		Normalize_VersionNumber = objRegEx.Replace (version, "$1.0")
	Else
		Normalize_VersionNumber = version
	End If

	objRegEx = Null
End Function

Class IEVersionInfo
	Public Version
	Public VersionName
End Class

Function Assist_CheckIEVersion(TargetVersion, OfficalVersion, OfficalVersionName)
	If OfficalVersionName = "*" Then
		Set Assist_CheckIEVersion = New IEVersionInfo
		Assist_CheckIEVersion.Version = TargetVersion
		Assist_CheckIEVersion.VersionName = ""
		Exit Function
	End If

	Dim normlizedVersion
	normlizedVersion = Normalize_VersionNumber(OfficalVersion)

	If TargetVersion <> normlizedVersion Then
		Set Assist_CheckIEVersion = Nothing
		Exit Function
	End If

	Set Assist_CheckIEVersion = New IEVersionInfo
	Assist_CheckIEVersion.Version = OfficalVersion
	Assist_CheckIEVersion.VersionName = OfficalVersionName
End Function


IEKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\"
Set WshShell = CreateObject("WScript.Shell")
On Error Resume Next
svcVersion = Null
svcVersion = WshShell.RegRead(IEKeyPath & "svcVersion")
version = Null
version = WshShell.RegRead(IEKeyPath & "Version")
On Error GOTO 0
Set WshShell = Nothing

' IE 10 or above
If Not IsNull(svcVersion) Then
	' Internet Explorer 10 Consumer Preview
	Set ieResult = Assist_CheckIEVersion(svcVersion, "10.0.8250.00000", "Internet Explorer 10 Consumer Preview")
	If Not ieResult Is Nothing Then
		WScript.echo ieResult.Version & " " & ieResult.VersionName
		WScript.Quit
	End If

	' Internet Explorer 10 Release Preview
	Set ieResult = Assist_CheckIEVersion(svcVersion, "10.0.8400.00000", "Internet Explorer 10 Release Preview")
	If Not ieResult Is Nothing Then
		WScript.echo ieResult.Version & " " & ieResult.VersionName
		WScript.Quit
	End If

	' Internet Explorer 10 for Windows 8
	Set ieResult = Assist_CheckIEVersion(svcVersion, "10.0.9200.16384", "Internet Explorer 10 for Windows 8")
	If Not ieResult Is Nothing Then
		WScript.echo ieResult.Version & " " & ieResult.VersionName
		WScript.Quit
	End If

	' Other
	Set ieResult = Assist_CheckIEVersion(svcVersion, "", "*")
	If Not ieResult Is Nothing Then
		WScript.echo ieResult.Version & " " & ieResult.VersionName
		WScript.Quit
	End If

	WScript.Quit
End If

' IE 4 ~ IE 9
'
' Internet Explorer 9 RTM
Set ieResult = Assist_CheckIEVersion(version, "9.0.8112.16421", "Internet Explorer 9 RTM")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 9 RC
Set ieResult = Assist_CheckIEVersion(version, "9.0.8080.16413", "Internet Explorer 9 RC")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 8 for Windows 7 and for Windows Server 2008 R2 (release)
Set ieResult = Assist_CheckIEVersion(version, "8.00.7600.16385", "Internet Explorer 8 for Windows 7 and for Windows Server 2008 R2 (release)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 8 for Windows 7 Beta
Set ieResult = Assist_CheckIEVersion(version, "8.00.7000.00000", "Internet Explorer 8 for Windows 7 Beta")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 8 for Windows XP, Windows Vista, Windows Server 2003 and Windows Server 2008
Set ieResult = Assist_CheckIEVersion(version, "8.00.6001.18702", "Internet Explorer 8 for Windows XP, Windows Vista, Windows Server 2003 and Windows Server 2008")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 8 RC1
Set ieResult = Assist_CheckIEVersion(version, "8.00.6001.18372", "Internet Explorer 8 RC1")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 8 Beta 2
Set ieResult = Assist_CheckIEVersion(version, "8.00.6001.18241", "Internet Explorer 8 Beta 2")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 8 Beta 1
Set ieResult = Assist_CheckIEVersion(version, "8.00.6001.17184", "Internet Explorer 8 Beta 1")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 7 for Windows Server 2008 and for Windows Vista SP1
Set ieResult = Assist_CheckIEVersion(version, "7.00.6001.1800", "Internet Explorer 7 for Windows Server 2008 and for Windows Vista SP1")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 7 for Windows XP SP2 x64
Set ieResult = Assist_CheckIEVersion(version, "7.00.6000.16441", "Internet Explorer 7 for Windows XP SP2 x64")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 7 for Windows Server 2003 SP2 x64
Set ieResult = Assist_CheckIEVersion(version, "7.00.6000.16441", "Internet Explorer 7 for Windows Server 2003 SP2 x64")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 7 for Windows Vista
Set ieResult = Assist_CheckIEVersion(version, "7.00.6000.16386", "Internet Explorer 7 for Windows Vista")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 7 for Windows XP and Windows Server 2003
Set ieResult = Assist_CheckIEVersion(version, "7.00.5730.1300", "Internet Explorer 7 for Windows XP and Windows Server 2003")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 7 for Windows XP and Windows Server 2003
Set ieResult = Assist_CheckIEVersion(version, "7.00.5730.1100", "Internet Explorer 7 for Windows XP and Windows Server 2003")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 6 SP2 for Windows Server 2003 SP1 and Windows XP x64
Set ieResult = Assist_CheckIEVersion(version, "6.00.3790.3959", "Internet Explorer 6 SP2 for Windows Server 2003 SP1 and Windows XP x64")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 6 for Windows Server 2003 SP1 and Windows XP x64
Set ieResult = Assist_CheckIEVersion(version, "6.00.3790.1830", "Internet Explorer 6 for Windows Server 2003 SP1 and Windows XP x64")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 6 for Windows Server 2003 (release)
Set ieResult = Assist_CheckIEVersion(version, "6.00.3790.0000", "Internet Explorer 6 for Windows Server 2003 (release)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 6 for Windows Server 2003 RC2
Set ieResult = Assist_CheckIEVersion(version, "6.00.3718.0000", "Internet Explorer 6 for Windows Server 2003 RC2")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 6 for Windows Server 2003 RC1
Set ieResult = Assist_CheckIEVersion(version, "6.00.3663.0000", "Internet Explorer 6 for Windows Server 2003 RC1")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 6 for Windows XP SP3
Set ieResult = Assist_CheckIEVersion(version, "6.00.2900.5512", "Internet Explorer 6 for Windows XP SP3")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 6 for Windows XP SP2
Set ieResult = Assist_CheckIEVersion(version, "6.00.2900.2180", "Internet Explorer 6 for Windows XP SP2")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 6 Service Pack 1 (Windows XP SP1)
Set ieResult = Assist_CheckIEVersion(version, "6.00.2800.1106", "Internet Explorer 6 Service Pack 1 (Windows XP SP1)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 6 (Windows XP)
Set ieResult = Assist_CheckIEVersion(version, "6.00.2600.0000", "Internet Explorer 6 (Windows XP)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 6 Public Preview (Beta) Refresh
Set ieResult = Assist_CheckIEVersion(version, "6.00.2479.0006", "Internet Explorer 6 Public Preview (Beta) Refresh")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 6 Public Preview (Beta)
Set ieResult = Assist_CheckIEVersion(version, "6.00.2462.0000", "Internet Explorer 6 Public Preview (Beta)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.5 Service Pack 2
Set ieResult = Assist_CheckIEVersion(version, "5.50.4807.2300", "Internet Explorer 5.5 Service Pack 2")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.5 Service Pack 1
Set ieResult = Assist_CheckIEVersion(version, "5.50.4522.1800", "Internet Explorer 5.5 Service Pack 1")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.5 Advanced Security Privacy Beta
Set ieResult = Assist_CheckIEVersion(version, "5.50.4308.2900", "Internet Explorer 5.5 Advanced Security Privacy Beta")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.5
Set ieResult = Assist_CheckIEVersion(version, "5.50.4134.0600", "Internet Explorer 5.5")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.5 for Windows Me (4.90.3000)
Set ieResult = Assist_CheckIEVersion(version, "5.50.4134.0100", "Internet Explorer 5.5 for Windows Me (4.90.3000)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.5 & Internet Tools Beta
Set ieResult = Assist_CheckIEVersion(version, "5.50.4030.2400", "Internet Explorer 5.5 & Internet Tools Beta")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.5 Developer Preview (Beta)
Set ieResult = Assist_CheckIEVersion(version, "5.50.3825.1300", "Internet Explorer 5.5 Developer Preview (Beta)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.01 SP4 (Windows 2000 SP4 only)
Set ieResult = Assist_CheckIEVersion(version, "5.00.3700.1000", "Internet Explorer 5.01 SP4 (Windows 2000 SP4 only)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.01 SP3 (Windows 2000 SP3 only)
Set ieResult = Assist_CheckIEVersion(version, "5.00.3502.1000", "Internet Explorer 5.01 SP3 (Windows 2000 SP3 only)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.01 SP2 (Windows 2000 SP2)
Set ieResult = Assist_CheckIEVersion(version, "5.00.3315.1000", "Internet Explorer 5.01 SP2 (Windows 2000 SP2)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.01 SP2 (Windows 95/98 and Windows NT 4.0)
Set ieResult = Assist_CheckIEVersion(version, "5.00.3314.2101", "Internet Explorer 5.01 SP2 (Windows 95/98 and Windows NT 4.0)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.01 SP1 (Windows 95/98 and Windows NT 4.0)
Set ieResult = Assist_CheckIEVersion(version, "5.00.3105.0106", "Internet Explorer 5.01 SP1 (Windows 95/98 and Windows NT 4.0)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.01 SP1 (Windows 2000 SP1)
Set ieResult = Assist_CheckIEVersion(version, "5.00.3103.1000", "Internet Explorer 5.01 SP1 (Windows 2000 SP1)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.01 (Windows 2000, build 5.00.2195)
Set ieResult = Assist_CheckIEVersion(version, "5.00.2920.0000", "Internet Explorer 5.01 (Windows 2000, build 5.00.2195)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.01 (Office 2000 SR-1)
Set ieResult = Assist_CheckIEVersion(version, "5.00.2919.6307", "Internet Explorer 5.01 (Office 2000 SR-1)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.01 (Windows 2000 RC2, build 5.00.2128)
Set ieResult = Assist_CheckIEVersion(version, "5.00.2919.3800", "Internet Explorer 5.01 (Windows 2000 RC2, build 5.00.2128)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.01 (Windows 2000 RC1, build 5.00.2072)
Set ieResult = Assist_CheckIEVersion(version, "5.00.2919.800", "Internet Explorer 5.01 (Windows 2000 RC1, build 5.00.2072)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5.01 (Windows 2000 Beta 3, build 5.00.2031)
Set ieResult = Assist_CheckIEVersion(version, "5.00.2516.1900", "Internet Explorer 5.01 (Windows 2000 Beta 3, build 5.00.2031)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5 (Windows 98 Second Edition)
Set ieResult = Assist_CheckIEVersion(version, "5.00.2614.3500", "Internet Explorer 5 (Windows 98 Second Edition)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5 (Office 2000)
Set ieResult = Assist_CheckIEVersion(version, "5.00.2314.1003", "Internet Explorer 5 (Office 2000)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5
Set ieResult = Assist_CheckIEVersion(version, "5.00.2014.0216", "Internet Explorer 5")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5 Beta (Beta 2)
Set ieResult = Assist_CheckIEVersion(version, "5.00.0910.1309", "Internet Explorer 5 Beta (Beta 2)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 5 Developer Preview (Beta 1)
Set ieResult = Assist_CheckIEVersion(version, "5.00.0518.10", "Internet Explorer 5 Developer Preview (Beta 1)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 4.01 Service Pack 2
Set ieResult = Assist_CheckIEVersion(version, "4.72.3612.1713", "Internet Explorer 4.01 Service Pack 2")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 4.01 Service Pack 1 (Windows 98)
Set ieResult = Assist_CheckIEVersion(version, "4.72.3110.8", "Internet Explorer 4.01 Service Pack 1 (Windows 98)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 4.01
Set ieResult = Assist_CheckIEVersion(version, "4.72.2106.8", "Internet Explorer 4.01")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 4.0
Set ieResult = Assist_CheckIEVersion(version, "4.71.1712.6", "Internet Explorer 4.0")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 4.0 Platform Preview 2.0 (PP2)
Set ieResult = Assist_CheckIEVersion(version, "4.71.1008.3", "Internet Explorer 4.0 Platform Preview 2.0 (PP2)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 4.0 Platform Preview 1.0 (PP1)
Set ieResult = Assist_CheckIEVersion(version, "4.71.544", "Internet Explorer 4.0 Platform Preview 1.0 (PP1)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 3.02 and 3.02a
Set ieResult = Assist_CheckIEVersion(version, "4.70.1300", "Internet Explorer 3.02 and 3.02a")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 3.01
Set ieResult = Assist_CheckIEVersion(version, "4.70.1215", "Internet Explorer 3.01")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 3.0 (Windows 95 OSR2)
Set ieResult = Assist_CheckIEVersion(version, "4.70.1158", "Internet Explorer 3.0 (Windows 95 OSR2)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 3.0
Set ieResult = Assist_CheckIEVersion(version, "4.70.1155", "Internet Explorer 3.0")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 2.0
Set ieResult = Assist_CheckIEVersion(version, "4.40.520", "Internet Explorer 2.0")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

' Internet Explorer 1.0 (Plus! for Windows 95)
Set ieResult = Assist_CheckIEVersion(version, "4.40.308", "Internet Explorer 1.0 (Plus! for Windows 95)")
If Not ieResult Is Nothing Then
	WScript.echo ieResult.Version & " " & ieResult.VersionName
	WScript.Quit
End If

'Other
WScript.echo "Can't determine IE version"
