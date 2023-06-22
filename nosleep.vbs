Set ws = CreateObject("Wscript.shell")

max = 6
cont = 0

Do While cont < max 'Max sleep in loop
	WScript.Sleep 299000 'Sleeps for 5 minutes
	ws.SendKeys "{CAPSLOCK}"
	cont = cont + 1
	
Loop

MsgBox("Finish!")

WScript.Quit()
	