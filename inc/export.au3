Func Export($sKeyCombo,$sFileName,$cType,$bSales = 0)
;set export filename
$export_filename = $sFileName

Do
WinActivate("AccountRight")
Send($sKeyCombo)
Until WinWaitActive("Export File","",5) <> 0

; identifiers/filters selected first. delete to obtain all records.
Send("{BS}")

; tab or csv
$winpos = WinGetPos("Export File")
MouseClick("left", $winpos[0] + 240, $winpos[1] + 76)
; down is tab, down down is csv.
If $cType == $TAB Then
	Send("{DOWN}{ENTER}")
	$ext = ".txt"
ElseIf $cType == $CSV Then
	Send("{DOWN}{DOWN}{ENTER}")
	$ext = ".csv"
EndIf

; include header record
MouseClick("left", $winpos[0] + 249, $winpos[1] + 112)
Send("{DOWN}{ENTER}")

If $bSales == 1 Then
; begin sales report config

; choose all sales
MouseClick("left", $winpos[0] + 240, $winpos[1] + 152)
Send("{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{ENTER}")

; choose date from
MouseClick("left", $winpos[0] + 220, $winpos[1] + 196)
Send("20/10/2010")

; choose date to
;MouseClick("left", $winpos[0] + 375, $winpos[1] + 196)
Send("{TAB}")
Send("{BS}")

; end sales report config
EndIf

Send("!o")
WinWaitActive("Export Data")
Send("!m")
Send("!e")
WinWaitActive("Save As")
; put filename
ControlSend("[CLASS:#32770]", "", "Edit1", $export_dir & $export_filename & $ext,1)
Send("!s")

; handle overwrite file prompt
If WinWaitActive("Confirm Save As","",2) Then
		Send("!y")
EndIf

; handle nothing to export
If WinWaitActive("","Nothing to export",10) Then
		Send("!o")
EndIf

EndFunc