Set WSHShell = CreateObject("WScript.Shell")

'WSHShell.Run "cmd\workTimeInput.wsf "& date & sTime & eTime & work
'WSHShell.Run "cmd\workTimeInput.wsf 20180619 9:00 18:00 ‚¢‚ë‚¢‚ë"


'Às“ú‚Ì9:00`18:00‚Ü‚Å‚Ì‹Î‘Ó‚ğ“ü—Í
WSHShell.Run "cmd\workTimeInput.wsf Replace(Left(Now(),10), "/", "") 9:00 18:00 ‚¢‚ë‚¢‚ë"
