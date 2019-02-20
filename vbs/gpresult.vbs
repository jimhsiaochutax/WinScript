Set WshShell = WScript.CreateObject("WScript.Shell")
Return = WshShell.Run("GPRESULT /H C:\GPReport.html", 0, false)
