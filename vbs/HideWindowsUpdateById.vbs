If Wscript.Arguments.Count = 0 Then
    WScript.Echo "Syntax: HideWindowsUpdateById.vbs [Update ID]" & vbCRLF & _
                 "Examples:" & vbCRLF & _
                 "  - Hide KB940157: HideWindowsUpdateById.vbs 2ba85467-deaf-44a1-a035-697742efab0f"
    WScript.Quit 1
End If

Dim updateId
updateId = WScript.Arguments(0)

Dim updateSession, updateSearcher
Set updateSession = CreateObject("Microsoft.Update.Session")
Set updateSearcher = updateSession.CreateUpdateSearcher()

Wscript.Stdout.Write "Searching for pending updates..." 
Dim searchResult
Set searchResult = updateSearcher.Search("UpdateID = '" & updateId & "'")

Dim update, index
WScript.Echo CStr(searchResult.Updates.Count) & " found."
For index = 0 To searchResult.Updates.Count - 1
    Set update = searchResult.Updates.Item(index)
    WScript.Echo "Hiding update: " & update.Title
    update.IsHidden = True
Next
