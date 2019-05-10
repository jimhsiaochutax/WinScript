@echo off
for /f "tokens=* skip=2 delims=" %%G in ('reg query "HKEY_CURRENT_USER\SOFTWARE\Policies\Google\Chrome\PopupsAllowedForUrls" /s') do (
call :sub "%%G")
for /f "tokens=* skip=2 delims=" %%G in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Google\Chrome\PopupsAllowedForUrls" /s') do (
call :sub "%%G")
::pause
goto :eof
 
:sub
  set _ie=%1
  set _ie=%_ie:"=%
  set _ie=%_ie: =%
  set _ie=%_ie:REG_SZ=% 
  set _ie=%_ie:1=%
  set _ie=%_ie:2=%
  set _ie=%_ie:3=%
  set _ie=%_ie:4=%
  set _ie=%_ie:5=%
  set _ie=%_ie:6=%
  set _ie=%_ie:7=%
  set _ie=%_ie:8=%
  set _ie=%_ie:9=%
  set _ie=%_ie:10=%

  echo %_ie%

