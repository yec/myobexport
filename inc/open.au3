Run($myob_executable)
WinWaitActive("Welcome to AccountRight")
Send("!o")
WinWaitActive("Open")
;input file to open
Sleep(1000)
Send("!n")
ControlSend("[CLASS:#32770]", "", "Edit1", $myob_fullpath,1)
Send("!o")
WinWaitActive("Sign-on")
Sleep(1000)
Send($myob_username,1)
Send("{TAB}")
Send($myob_password,1)
Send("!m")
Send("!o")
WinWaitActive("AccountRight Premier")
;Sleep(1000)