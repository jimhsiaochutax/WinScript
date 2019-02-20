@echo off


for /F "tokens=1-3 delims=: " %%i in ('echo %TIME%') do ( 
   set HH=%%i
   set MM=%%j
   set SS=%%k
)

::EQU - equal
::NEQ - not equal
::LSS - less than
::LEQ - less than or equal
::GTR - greater than
::GEQ - greater than or equal
set PM=0
if %HH% GTR 12 set /a HH-=12 & set PM=1

if "%PM%" == "0" (
  echo �{�b�ɶ�:	�W�� %HH%:%MM% > msg.txt
)

if "%PM%" == "1" if %HH% LSS 6 (
  echo �{�b�ɶ�:	�U�� %HH%:%MM% > msg.txt
)

if "%PM%" == "1" if %HH% GEQ 6 (
  echo �{�b�ɶ�:	�ߤW %HH%:%MM% > msg.txt
)

if "%PM%" == "1" if %HH% GEQ 10 (
  echo( >> msg.txt
  echo ������ ������ �{���� �~���F�C >> msg.txt
  echo �������� �������� �{���� ��ı�F�C >> msg.txt
)

if "%PM%" == "0" if %HH% LSS 6 (
  echo( >> msg.txt
  echo ������ ������ �{���� �~���F�C >> msg.txt
  echo �������� �������� �{���� ��ı�F�C >> msg.txt
)

msg * /TIME:10 < msg.txt


::ipconfig /release "�ϰ�s�u 2"
::ipconfig /renew "�ϰ�s�u 2"
::timeout /t 60
