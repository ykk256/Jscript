Set WSHShell = CreateObject("WScript.Shell")

'WSHShell.Run "cmd\workTimeInput.wsf "& date & sTime & eTime & work
'WSHShell.Run "cmd\workTimeInput.wsf 20180619 9:00 18:00 いろいろ"


'実行日の9:00〜18:00までの勤怠を入力
WSHShell.Run "cmd\workTimeInput.wsf Replace(Left(Now(),10), "/", "") 9:00 18:00 いろいろ"
