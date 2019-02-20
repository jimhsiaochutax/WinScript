
dim NIC1, Nic, StrIP
Set NIC1 = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

For Each Nic in NIC1
  if Nic.IPEnabled then
    StrIP = Nic.IPAddress(i)
    MsgBox "IP Address:  " & StrIP & vbNewLine
  end if
Next

MsgBox "IP Address:  " & StrIP & vbNewLine
