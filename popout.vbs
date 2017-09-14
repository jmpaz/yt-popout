Dim objHTA
Dim clipboard
Dim WshShell
set objHTA = createobject("htmlfile")
clipboard = objHTA.parentwindow.clipboarddata.getdata("text")
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"" --app=https://www.youtube.com/embed/" & clipboard & """"