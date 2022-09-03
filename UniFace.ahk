#Requires Autohotkey v1.1.33+
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

:*:]uni:: ; Radio button with choice of Family code for Peoplesoft/Excel
CoordMode, Mouse, Screen
GUI, Destroy
FormatTime, CurrentDateTime,, yyyy-MM-dd
MouseGetPos, xpos, ypos
xpos := xpos - 80
ypos := ypos - 80
GUI, Add, Radio, altsubmit gCheckuni vRadioGroupuni, ¯\_(?)_/¯ (Shrug)
GUI, Add, Radio, altsubmit gCheckuni, (° ???°)?n? (Flip)
GUI, Add, Radio, altsubmit gCheckuni, (? '`-'´)? (Fight)
GUI, Add, Radio, altsubmit gCheckuni, ?(-´-´?) (Boxer)
GUI, Add, Radio, altsubmit gCheckuni, [¬º-°]¬ (Zombi)
GUI, Add, Radio, altsubmit gCheckuni, (????)? (You da man)
GUI, Add, Radio, altsubmit gCheckuni, (?°?°)?? ??? (Table Flip)
GUI, Add, Radio, altsubmit gCheckuni, (n`-')????.*???  (PFM)

GUI, Show, x%xpos% y%ypos%
Return

Checkuni:
GUI, submit, nohide
If (RadioGroupuni = 1)
{
	RadioGroupuni := "¯\_(?)_/¯"
}
If (RadioGroupuni = 2)
{
	RadioGroupuni := "(° ???°)?n?"
}
If (RadioGroupuni = 3)
{
	RadioGroupuni := "(? '`-'´)?"
}
If (RadioGroupuni = 4)
{
	RadioGroupuni := "?(-´-´?)"
}
If (RadioGroupuni = 5)
{
	RadioGroupuni := "[¬º-°]¬"
}
If (RadioGroupuni = 6)
{
	RadioGroupuni := "(????)?"
}
If (RadioGroupuni = 7)
{
	RadioGroupuni := "(?°?°)?? ???"
}
If (RadioGroupuni = 8)
{
	RadioGroupuni := "(n`-')????.*??? "
}

GUI, Destroy
Send, %RadioGroupuni%

Loop, 0xFF
{
	Key := Format("VK{:02X}",A_Index)
	IF GetKeyState(Key)
		Send, {%Key% Up}
}

Return