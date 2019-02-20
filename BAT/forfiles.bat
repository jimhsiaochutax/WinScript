
::@ECHO OFF

for /r %%I in (*.png) do convert.exe %%~nxI -crop 1872x1314+192+0 out\%%~nxI


