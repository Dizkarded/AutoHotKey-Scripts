#Requires Autohotkey v1.1.09+
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance, Force
/*
	General information:	The "!" is used as a symbol for {Alt}
	The "#" is used as a symbol for {Windows Key}
	The "^" is used as a symbol for {Control}
	The "+" is used as a symbol for {Shift}
	Anything after ";" is merely a comment and will not be executed in the script
*/



InputBox, NAMEDEPT, Name to HCPCS, Enter the Name / Department.?`n`n[Hit {ESC} at anytime to exit.],,450,175,,,,,Mike Behnami / MOR
If ErrorLevel
	ExitApp

InputBox, HCPCS, Name to HCPCS, Enter the HCPCS Code.`n`n[Hit {ESC} at anytime to exit.],,450,175,,,,,N/A
If ErrorLevel
	ExitApp


F1::
Gosub, NH
Return

NH:
Send, {Tab 3}
Sleep 550
Send, %NAMEDEPT%{Tab 13}%HCPCS%
Sleep 550
Send, {Alt Down}1{Alt Up}
Sleep 1500
Send, {Alt Down}3{Alt Up}
Sleep 1500
Send, {Alt Down}t{Alt Up}

Gosub, NAMEHCPCSAllKeysUp

Return

NAMEHCPCSAllKeysUp:

Loop, 0xFF
{
	Key := Format("VK{:02X}",A_Index)
	IF GetKeyState(Key)
		Send, {%Key% Up}
}

Return

^esc:: ; (Control + Escape) Reload script
ExitApp
Return