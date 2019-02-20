
Dim oShell
Dim oStations

Set oShell = WScript.CreateObject ("WScript.Shell")

oStations = Array(
"pc0",
"pc1"
)


For Each Station In oStations
    oShell.run "rem shutdown /r /f /t 0 /m \\" & Station
Next Station

'oShell.run "cmd /K CD C:\ & Dir"
Set oShell = Nothing

