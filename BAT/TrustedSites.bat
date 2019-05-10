@echo off
:: TrustedSites.cmd
:: Display Trusted sites and Intranet sites
:: HKCU\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMapKey
:: HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\
for /f "tokens=* skip=3 delims=" %%G in ('reg query "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains" /s') do (
call :sub "%%G")
::pause
goto :eof
 
:sub
  set _ie=%1
  set _ie=%_ie:"=%
  set _ie=%_ie:REG_SZ= - % 
  set _ie=%_ie:REG_DWORD=% 
  set _ie=%_ie:HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\=% 
  set _ie=%_ie:  0x0= My Computer%
  set _ie=%_ie:  0x1= Local Intranet Zone%
  set _ie=%_ie:  0x2= Trusted sites Zone%
  set _ie=%_ie:  0x3= Internet Zone%
  set _ie=%_ie:  0x4= Restricted Sites Zone%

  echo %_ie%

