Dim obj
Dim clipboard
Dim WshShell
Dim vidID
set obj = createobject("htmlfile")
clipboard = obj.parentwindow.clipboarddata.getdata("text")
vidID = Right(clipboard, Len(clipboard) - InStrRev(clipboard, "="))

If Len(vidID) = 11 then
	Set WshShell = WScript.CreateObject("WScript.Shell")
	WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"" --app=https://www.youtube.com/embed/" & vidID & """"
ElseIf Len(vidID) = 28 or Len(vidID) <= 5 then
   'fix for youtu.be format:
	vidID = Right(clipboard, Len(clipboard) - InStrRev(clipboard, "/"))
	WScript.Echo(Len(vidID))
	WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"" --app=https://www.youtube.com/embed/" & vidID & """"
   'this section currently not working
Else 
	WScript.Echo "failure. timestamps currently only working in youtu.be format. vidID length: " & Len(vidID)
End If