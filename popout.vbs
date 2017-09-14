Dim obj
Dim clipboard
Dim WshShell
Dim baseURL
Dim vidID
set obj = createobject("htmlfile")
clipboard = obj.parentwindow.clipboarddata.getdata("text")
baseURL = ' to make sure this is a youtube link
vidID = Right(clipboard, Len(clipboard) - InStrRev(clipboard, "=")) ' measures the length of everything after =

If  baseURL  'pseudo: = youtube.com OR youtu.be then ||| note: this will require a rewrite of the code below but is much more extensible
	If Len(vidID) = 11 then
		Set WshShell = WScript.CreateObject("WScript.Shell")
		WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"" --app=https://www.youtube.com/embed/" & vidID & """"
	ElseIf Len(vidID) = 28 ' this means it's a standard youtu.be link (there was no = , so it )
	'fix for youtu.be format:
		vidID = Right(clipboard, Len(clipboard) - InStrRev(clipboard, "/"))
		WScript.Echo(Len(vidID)) 'this line for testing, will be removed
		WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"" --app=https://www.youtube.com/embed/" & vidID & """"
	'this section currently not working
	Else
		vidID = Right(clipboard, Len(clipboard) - InStrRev(clipboard, "?"))
	Else 
		WScript.Echo "failure. timestamps currently only working in youtu.be format. vidID length: " & Len(vidID)
	End If
Else