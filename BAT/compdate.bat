
@ECHO OFF

::FOR %%F IN (*.LOG) DO 7z.exe A %%F.7z %%F & DEL %%F

forfiles /d -1 /m "*.LOG" /c "cmd /c echo @file"

