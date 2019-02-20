wmic qfe get hotfixid | find "KB4093108" > kb.txt
wmic qfe | find "KB4093108" >> kb.txt
wmic qfe get hotfixid | find "KB4093118" >> kb.txt
wmic qfe | find "KB4093118" >> kb.txt