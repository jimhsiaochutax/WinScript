
@ECHO OFF

forfiles /d -1 /m *.LOG /c "cmd /Q /C for %%I in (@file) do 7z.exe A %%~I.7z %%~I & DEL %%~I"

