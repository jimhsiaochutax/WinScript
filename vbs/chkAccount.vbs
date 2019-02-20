
strComputer = "."
'strComputer = "10611PC010"

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colAccounts = objWMIService.ExecQuery _
    ("Select * From Win32_UserAccount Where LocalAccount = TRUE")
'    ("Select * From Win32_Group Where LocalAccount = TRUE And SID = 'S-1-5-32-544'")

For Each objAccount in colAccounts
    Wscript.Echo objAccount.Domain
	Wscript.Echo objAccount.Name
	Wscript.Echo objAccount.Disabled
Next

'在這裡我們執行的動作是針對符合下列條件電腦上的群組發出 WMI 查詢：
'LocalAccount 屬性的值為 True。這是查詢中很重要的一個部分，因為它限制 WMI 只能查詢本機帳戶，如果沒有設定，那麼 WMI 會查詢 Active Directory 裡每一個群組。
'SID 屬性的值為 S-1-5-32-544。SID (安全識別碼) 是一組獨一無二的號碼，可供作業系統識別帳戶。因此您可以變更本機管理員帳戶的名稱，無須擔心本機管理員就此無法進行存取。大部分的情況下，作業系統會採 SID 而忽略使用帳戶名稱。而且這個值永遠都不會更改。
'所以說 S-1-5-32-544 值相當重要囉？沒錯，這個 SID 號碼碰巧「很出名」，而且永遠參考本機管理員群組。也就是說，如果您找到的本機帳號 SID 為 S-1-5-32-544，賓果！您就找到了本機管理員群組。就是這麼簡單！
'剩下來的指令碼很簡單，只要設定一個 For Each 迴圈，以 SID S-1-5-32-544 跑過所有的本機群組並傳回每個群組的名稱即可。由於每台電腦上的 SID 都不一樣，而且 S-1-5-32-544 指的一定是本機管理員群組，您蒐集到的結果正是本機管理員群組，僅此一個，別無他物。這樣一來，不管您的作業系統採用何種語言，不管您的本機管理員群組名稱為何，都找得到。
'但是這個指令碼的確有其缺陷，前面曾說過，它只適用 Windows XP 和 Windows Server 2003。Windows 2000 的使用者別擔心，咱們可沒忘了您。下面是改寫過的指令碼，適用 Windows 2000、Windows XP、Windows Server 2003，甚至 Windows NT 4.0：
'strComputer = "atl-fs-01"
'
'Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
'
'Set colAccounts = objWMIService.ExecQuery _
'    ("Select * From Win32_Group Where Domain = '" & strComputer & "' AND SID = 'S-1-5-32-544'")
'
'For Each objAccount in colAccounts
'    Wscript.Echo objAccount.Name
'Next

