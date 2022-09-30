
Function DateString(dDate)
    DateString = Year(dDate)& right("_" & Month(dDate),2) & right("_" & Day(dDate),2) & right("_" & Hour(dDate),2) & right("0" & Minute(dDate),2) & right("0" & second(dDate),2)
End Function
sql="BACKUP DATABASE vhotelhdd TO DISK = 'D:\vhotelhdd" &DateString(now())& ".bak'"

Wscript.echo sql

Function GetDownloadsPath() 
   Dim sDesktopPath
Set objShell = Wscript.CreateObject("Wscript.Shell")
GetDownloadsPath = objShell.SpecialFolders("Desktop")
End Function

WScript.Echo Replace(GetDownloadsPath(),"Desktop","Downloads")