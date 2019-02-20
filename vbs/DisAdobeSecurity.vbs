'
'
' Create a WSH Shell object:
Set wshShell = CreateObject( "WScript.Shell" )
'
' Create a new key:
'wshShell.RegWrite "HKCU\TestKey\", ""

' Create a new DWORD value:
'wshShell.RegWrite "HKCU\TestKey\DWordTestValue", 1, "REG_DWORD"

' Create a new subkey and a string value in that new subkey:
'wshShell.RegWrite "HKCU\TestKey\SubKey\StringTestValue", "Test", "REG_SZ"

' Read the values we just created:
'WScript.Echo "HKCU\TestKey\DWordTestValue = " _
'           & wshShell.RegRead( "HKCU\TestKey\DWordTestValue" )
'WScript.Echo "HKCU\TestKey\SubKey\StringTestValue = """ _
'           & wshShell.RegRead( "HKCU\TestKey\SubKey\StringTestValue" ) & """"

' Delete the subkey and key and the values they contain:
'wshShell.RegDelete "HKCU\TestKey\SubKey\"
'wshShell.RegDelete "HKCU\TestKey\"

' Note: Since the WSH Shell has no Enumeration functionality, you cannot
'       use the WSH Shell object to delete an entire "tree" unless you
'       know the exact name of every subkey.
'       If you don't, use the WMI StdRegProv instead.

'msgbox wshShell.RegRead( "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Acrobat Reader\11.0\FeatureLockDown\cDefaultExecMenuItems\tWhiteList" ) 
If(wshShell.RegRead( "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Acrobat Reader\9.0\FeatureLockDown\cDefaultExecMenuItems\tWhiteList" ) <> "") Then
  wshShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Acrobat Reader\9.0\FeatureLockDown\bProtectedMode", 0, "REG_DWORD" 
End If
If(wshShell.RegRead( "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Acrobat Reader\10.0\FeatureLockDown\cDefaultExecMenuItems\tWhiteList" ) <> "") Then
  wshShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Acrobat Reader\10.0\FeatureLockDown\bProtectedMode", 0, "REG_DWORD" 
End If
If(wshShell.RegRead( "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Acrobat Reader\11.0\FeatureLockDown\cDefaultExecMenuItems\tWhiteList" ) <> "") Then
  wshShell.RegWrite "HKLM\SOFTWARE\Policies\Adobe\Acrobat Reader\11.0\FeatureLockDown\bProtectedMode", 0, "REG_DWORD" 
End If
If(wshShell.RegRead( "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown\cDefaultExecMenuItems\tWhiteList" ) <> "") Then
  wshShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown\bProtectedMode", 0, "REG_DWORD" 
End If
'msgbox wshShell.RegRead( "HKLM\Software\Microsoft\Windows NT\CurrentVersion\winlogon\AutoAdminLogon" )

' Release the object
Set wshShell = Nothing
