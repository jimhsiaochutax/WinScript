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
  echo 現在時間:	上午 %HH%:%MM% > msg.txt
)

if "%PM%" == "1" if %HH% LSS 6 (
  echo 現在時間:	下午 %HH%:%MM% > msg.txt
)

if "%PM%" == "1" if %HH% GEQ 6 (
  echo 現在時間:	晚上 %HH%:%MM% > msg.txt
)

if "%PM%" == "1" if %HH% GEQ 10 (
  echo( >> msg.txt
  echo ㄒㄧˇ ㄗㄠˇ ㄌㄜ˙ 洗澡了。 >> msg.txt
  echo ㄕㄨㄟˋ ㄐㄧㄠˋ ㄌㄜ˙ 睡覺了。 >> msg.txt
)

if "%PM%" == "0" if %HH% LSS 6 (
  echo( >> msg.txt
  echo ㄒㄧˇ ㄗㄠˇ ㄌㄜ˙ 洗澡了。 >> msg.txt
  echo ㄕㄨㄟˋ ㄐㄧㄠˋ ㄌㄜ˙ 睡覺了。 >> msg.txt
)

msg * /TIME:10 < msg.txt


::ipconfig /release "區域連線 2"
::ipconfig /renew "區域連線 2"
::timeout /t 60
