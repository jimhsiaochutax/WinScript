If Wscript.Arguments.Count = 0 Then
    WScript.Echo "Syntax: HideWindowsUpdate.vbs [Hotfix Article ID]" & vbCRLF & _
                 "Examples:" & vbCRLF & _
                 "  - Hide KB940157: HideWindowsUpdate.vbs 940157"
    WScript.Quit 1
End If

Dim hotfixId
hotfixId = WScript.Arguments(0)

Dim updateSession, updateSearcher
Set updateSession = CreateObject("Microsoft.Update.Session")
Set updateSearcher = updateSession.CreateUpdateSearcher()

Wscript.Stdout.Write "Searching for pending updates..." 
Dim searchResult
Set searchResult = updateSearcher.Search("IsInstalled=0")

Dim update, kbArticleId, index, index2
WScript.Echo CStr(searchResult.Updates.Count) & " found."
For index = 0 To searchResult.Updates.Count - 1
    Set update = searchResult.Updates.Item(index)
    For index2 = 0 To update.KBArticleIDs.Count - 1
        kbArticleId = update.KBArticleIDs(index2)
        If kbArticleId = hotfixId Then
            WScript.Echo "Hiding update: " & update.Title
            update.IsHidden = True
        End If        
    Next
Next
