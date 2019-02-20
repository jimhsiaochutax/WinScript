C:\Windows\System32\rundll32.exe printui.dll, PrintUIEntry /y /n "印表機名稱"


新增網路印表機：
　rundll32 printui.dll,PrintUIEntry /in /q /n \\(伺服器名稱)\(印表機分享名稱)
將網路印表機設定為「預設印表機」：
　rundll32 printui.dll,PrintUIEntry /y /q /n \\(伺服器名稱)\(印表機分享名稱)
刪除網路印表機：
　rundll32 printui.dll,PrintUIEntry /dn /n \\(伺服器名稱)\(印表機分享名稱)
資料來源
rundll32 printui.dll,PrintUIEntry /dn /n "\\printer\CanoniR3235(2F辦公室)"
rundll32 printui.dll,PrintUIEntry /dn /n "\\printer\CanoniR3235(總務處)"

rundll32 printui.dll,PrintUIEntry /dn /n "\\printer\Canon iR3235/iR3245 UFR II(2F辦公室)"


@echo off
::if not exist \\filesrv-vm\gpowork\InstallPrinter-test.txt goto End  <==判斷目錄與檔案是否存在，若不存在就跳離批次檔
if not exist \\files\misc$\設定印表機\add-printer-rs.txt goto End
rem <==判斷目錄與檔案是否存在，若不存在就跳離批次檔

@type \\files\misc$\設定印表機\add-printer-rs.txt |find/I "[%COMPUTERNAME%][%USERNAME%]" > nul 2>&1
rem <==查詢 txt 檔內容是否有對應的電腦名稱與使用者名稱，若有則不再執行新增印表機動作，若沒有則開始新增印表機的動作 
if %errorlevel% == 0 goto End

rundll32 printui.dll,PrintUIEntry /in /n "\\printer\SHARP總務處"
rundll32 printui.dll,PrintUIEntry /in /n "\\printer\SHARP辦公室"
rundll32 printui.dll,PrintUIEntry /y /n "\\printer\SHARP辦公室"

rundll32 printui.dll,PrintUIEntry /dn /q /n "\\printer\Canon iR3235/iR3245 UFR II(2F辦公室)"
rundll32 printui.dll,PrintUIEntry /dn /q /n "\\printer\Canon iR3235/iR3245 UFR II(總務處)"
rundll32 printui.dll,PrintUIEntry /dn /q /n "\\printer\CanoniR3235(2F辦公室)"
rundll32 printui.dll,PrintUIEntry /dn /q /n "\\printer\CanoniR3235(總務處)"

echo [%COMPUTERNAME%][%USERNAME%] >> \\files\misc$\設定印表機\add-printer-rs.txt
rem <==新增完印表機則將電腦名稱與使用者名稱寫入txt檔案內

exit
:End


