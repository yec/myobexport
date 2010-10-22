; Export MYOB

; allow only one script to run at a time
#Include <Misc.au3>
if _Singleton("myobexport",1) = 0 Then
    Msgbox(16,"Error","MYOB Export is already running. You should shutdown the process in the task manager.")
    Exit
EndIf

$ini_file = "config.ini"

$myob_fullpath = IniRead ( $ini_file, "myob", "fullpath", "c:\premier19\clearwtr.myo" )
$myob_username = IniRead ( $ini_file, "myob", "username", "administrator" )
$myob_password = IniRead ( $ini_file, "myob", "password", "" )
$myob_executable = IniRead ( $ini_file, "myob", "executable", "c:\premier19\myobp.exe" )

$export_dir = IniRead ( $ini_file, "export", "dir", "c:\myobwexports\" )

; create export directory if not existing yet
if not fileexists($export_dir) then
	dircreate($export_dir)
endif

Opt("SendKeyDelay",10)

$keepopen = True
; open myob if not open yet
if WinExists("AccountRight") = False Then
	#include "inc/open.au3"
	$keepopen = False
EndIf

WinActivate("AccountRight")

#include "inc/export.au3"

Export("{ALTDOWN}f{ALTUP}ecc","customers.csv")
Export("{ALTDOWN}f{ALTUP}ecs","suppliers.csv")
Export("{ALTDOWN}f{ALTUP}ei","items.csv")
Export("{ALTDOWN}f{ALTUP}ess{RIGHT}i","itemsales.csv",1)
Export("{ALTDOWN}f{ALTUP}ess{RIGHT}s","servicesales.csv",1)

if $keepopen = False Then
	WinClose("AccountRight Premier")
EndIf

Exit