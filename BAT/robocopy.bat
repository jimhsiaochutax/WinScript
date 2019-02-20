@echo on
robocopy Source Destination /E /ZB /X /PURGE /COPYALL /TEE /DCOPY:T /LOG:C:\ofcscan.txt
