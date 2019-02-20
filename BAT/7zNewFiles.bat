
for /F "tokens=1 " %%i in ('date /t') do (
   set STRBUF=%%i
)

for /F "tokens=1-4 delims=/ " %%i in ('echo %STRBUF%') do (
   set YEAR=%%i
   set MONTH=%%j
   set DAY=%%k
)

echo %YEAR%/%MONTH%/%DAY%

   if %MONTH:~0,1%==0 set MONTH=%MONTH:~1%
   set /a MONTH-=1 && if %MONTH%==0 set MONTH=12&set /a YEAR-=1
   set MONTH=0%MONTH%
   set MONTH=%MONTH:~-2%

echo %YEAR%/%MONTH%/%DAY%

forfiles /p D: /s /c "cmd /c if @isdir==FALSE echo @relpath " /d +%YEAR%/%MONTH%/01 > list.txt
7za a -spf %YEAR%%MONTH%.7z -ir@list.txt

::setlocal enableDelayedExpansion
::for /f "delims=/ tokens=1,3" %%a in ("%date%") do set month=%%a& set year=%%b

::if %month:~0,1%==0 set month=%month:~1%
::set /a month-=1 && if !month!==0 set month=12&set /a year-=1

::if not !month!==2 (
::    set /a last_day="31 - (month - 1) %% 7 %% 2"
::) else (
::    set /a y4="year %% 4" & if !y4!==0 (
::        set /a y100="year %% 100" & if not !y100!==0 set is_leap_year=1
::        set /a y400="year %% 400" & if !y400!==0 set is_leap_year=1
::    )
::    if "!is_leap_year!"=="1" (set last_day=29) else set last_day=28
::)

::set month=0!month!
::start "Testing" "c:\Program Files\app-cmd\bin\admincmd\imagelist" ^
::    -d !month:~-2!/01/!year! -e !month:~-2!/!last_day!/!year! > c:\test.txt

