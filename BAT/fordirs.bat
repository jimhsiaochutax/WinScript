@ECHO OFF

for /D %%I in (*) do 7z.exe A %%I.7z %%I & RD /S /Q %%I

