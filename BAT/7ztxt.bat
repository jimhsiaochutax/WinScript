
@ECHO OFF

FOR %%F IN (*.TXT) DO 7z.exe A %%F.7z %%F & DEL %%F

