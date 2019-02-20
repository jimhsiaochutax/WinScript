
@ECHO OFF

DIR D:\Backup\Web /B /OD > LIST.TXT
FOR /F "DELIMS=?" %%F IN (LIST.TXT) DO (SET FNAME=%%F)
REM ECHO %FNAME%

"C:\Program Files\7-Zip\7z.exe" X -OD:\Web -Y "D:\Backup\Web\%FNAME%"

REM 取得日期
REM FOR /F "tokens=1-3 delims=/ " %%a IN ("%date%") DO (
REM FOR /F "tokens=1-3 delims=/ " %%a IN ('date /T') DO (
REM SET _MyDate=%%a%%b%%c
REM )

REM 顯示去掉分隔符號後的結果
REM echo %_MyDate%

REM 取得時間
REM FOR /F "tokens=1-4 delims=:." %%a IN ("%time%") DO (
REM FOR /F "tokens=1-3 delims=: " %%a IN ('time /T') DO (
REM SET _MyTime=%%a%%b%%c
REM )

REM 顯示去掉分隔符號後的結果
REM echo %_MyTime%

COPY /Y chutaxConnDb.inc.php D:\Web\Includes\

