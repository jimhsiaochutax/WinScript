C:\Windows\System32\rundll32.exe printui.dll, PrintUIEntry /y /n "�L����W��"


�s�W�����L����G
�@rundll32 printui.dll,PrintUIEntry /in /q /n \\(���A���W��)\(�L������ɦW��)
�N�����L����]�w���u�w�]�L����v�G
�@rundll32 printui.dll,PrintUIEntry /y /q /n \\(���A���W��)\(�L������ɦW��)
�R�������L����G
�@rundll32 printui.dll,PrintUIEntry /dn /n \\(���A���W��)\(�L������ɦW��)
��ƨӷ�
rundll32 printui.dll,PrintUIEntry /dn /n "\\printer\CanoniR3235(2F�줽��)"
rundll32 printui.dll,PrintUIEntry /dn /n "\\printer\CanoniR3235(�`�ȳB)"

rundll32 printui.dll,PrintUIEntry /dn /n "\\printer\Canon iR3235/iR3245 UFR II(2F�줽��)"


@echo off
::if not exist \\filesrv-vm\gpowork\InstallPrinter-test.txt goto End  <==�P�_�ؿ��P�ɮ׬O�_�s�b�A�Y���s�b�N�����妸��
if not exist \\files\misc$\�]�w�L���\add-printer-rs.txt goto End
rem <==�P�_�ؿ��P�ɮ׬O�_�s�b�A�Y���s�b�N�����妸��

@type \\files\misc$\�]�w�L���\add-printer-rs.txt |find/I "[%COMPUTERNAME%][%USERNAME%]" > nul 2>&1
rem <==�d�� txt �ɤ��e�O�_���������q���W�ٻP�ϥΪ̦W�١A�Y���h���A����s�W�L����ʧ@�A�Y�S���h�}�l�s�W�L������ʧ@ 
if %errorlevel% == 0 goto End

rundll32 printui.dll,PrintUIEntry /in /n "\\printer\SHARP�`�ȳB"
rundll32 printui.dll,PrintUIEntry /in /n "\\printer\SHARP�줽��"
rundll32 printui.dll,PrintUIEntry /y /n "\\printer\SHARP�줽��"

rundll32 printui.dll,PrintUIEntry /dn /q /n "\\printer\Canon iR3235/iR3245 UFR II(2F�줽��)"
rundll32 printui.dll,PrintUIEntry /dn /q /n "\\printer\Canon iR3235/iR3245 UFR II(�`�ȳB)"
rundll32 printui.dll,PrintUIEntry /dn /q /n "\\printer\CanoniR3235(2F�줽��)"
rundll32 printui.dll,PrintUIEntry /dn /q /n "\\printer\CanoniR3235(�`�ȳB)"

echo [%COMPUTERNAME%][%USERNAME%] >> \\files\misc$\�]�w�L���\add-printer-rs.txt
rem <==�s�W���L����h�N�q���W�ٻP�ϥΪ̦W�ټg�Jtxt�ɮפ�

exit
:End


