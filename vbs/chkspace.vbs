'Const ForWriting = 2
'Const ForAppending = 8
Const HARD_DISK = 3 

' file io
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set objFile = objFSO.OpenTextFile("C:\FSO\ScriptLog.txt", ForWriting)

' check computer name
Set wshShell = WScript.CreateObject( "WScript.Shell" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
'WScript.Echo "Computer Name: " & strComputerName
Set objFile = objFSO.CreateTextFile("\\nas\chkspace\" & strComputerName & ".txt", True)
objFile.Writeline ("=== " & strComputerName & " ===")

' check disk space 
strComputer = "." 
Set objWMIService = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
 
Set colDisks = objWMIService.ExecQuery _ 
    ("Select * from Win32_LogicalDisk Where DriveType = " & HARD_DISK & "") 
 
For Each objDisk in colDisks 
    'Wscript.Echo "DeviceID: "& vbTab &  objDisk.DeviceID        
    'Wscript.Echo "Free Disk Space: "& vbTab & objDisk.FreeSpace 
	'if objDisk.DeviceID = "D:" and objDisk.FreeSpace < 3221225472 then
	'Set objFile = objFSO.CreateTextFile("\\nas\chkspace\" & strComputerName & ".txt", True)
'objFile.Writeline ("=== " & strComputerName & " ===")
	objFile.Writeline (objDisk.DeviceID & vbTab & FormatNumber((CDbl(objDisk.FreeSpace)/1024/1024/1024)) & "G")
'objFile.Close
	'end if
Next 

'objFile.Write Now
'objFile.Writeline ("This is line 1.")
objFile.Close
