Dim objHTA
Dim clipboard
Dim WshShell
Dim vidID
set objHTA = createobject("htmlfile")
clipboard = objHTA.parentwindow.clipboarddata.getdata("text")
vidID = Right(clipboard, Len(clipboard) - InStrRev(clipboard, "="))
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"" --app=https://www.youtube.com/embed/" & vidID & """"