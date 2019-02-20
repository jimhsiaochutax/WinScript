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
msgbox wshShell.RegRead( "HKLM\Software\Microsoft\Windows NT\CurrentVersion\winlogon\AutoAdminLogon" )

' Delete the subkey and key and the values they contain:
'wshShell.RegDelete "HKCU\TestKey\SubKey\"
'wshShell.RegDelete "HKCU\TestKey\"

' Note: Since the WSH Shell has no Enumeration functionality, you cannot
'       use the WSH Shell object to delete an entire "tree" unless you
'       know the exact name of every subkey.
'       If you don't, use the WMI StdRegProv instead.

' Release the object
Set wshShell = Nothing
