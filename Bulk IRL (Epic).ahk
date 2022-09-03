#Requires Autohotkey v1.1.33+
; Created by Scott Stutzman - scott-stutzman@uiowa.edu
; Version Number  2021.03.26.15.43.26

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance, Force ; Allows only one instance of scirpt to run at a time.

/*
	General information:	The "!" is used as a symbol for {Alt}
	The "#" is used as a symbol for {Windows Key}
	The "^" is used as a symbol for {Control}
	The "+" is used as a symbol for {Shift}
	Anything after ";" is merely a comment and will not be executed in the script
*/
xcl := ComObjActive("Excel.Application") ; Variable "xcl" is now set to the last Excel work sheet that was active
InputBox, Row, What row to start on,What row would you like to start on?
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A
FormatTime, CurrentDateTime,, MM/dd/yyyy

; !`::
Numpad0::
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A
xcl.Range("A"Row).Interior.ColorIndex := 4
Clipboard := ColumnA
Send,  ^1
Sleep 500
Send, ^v
Sleep 500
Send, {Enter}
Sleep 300
Send, {Enter}
Sleep 300
Send, {Enter}
;return
Sleep 300
; !1::
;F2::
Send, {Alt Down}{Down 4}{Alt Up}
Sleep 300
Send, {Tab 2}{Down 4}IRL{Tab}IRL{Tab}
Sleep 300
Send, {Shift Down}{Tab 12}{Shift Up}
Sleep 300
Send, {Enter}IRL{Tab}{Enter}
Sleep 300
SendInput %CurrentDateTime%
Send, {Tab}{Enter}
Sleep 300
Send, y{Tab}
SLeep 300
Send, !a
;Send, {Tab 2}{Down 4}IRL{Tab}IRL{Tab}
;Sleep 300
;Send, {Tab}
SLeep 300
Send, !a
xcl.Range("A"Row).Interior.ColorIndex := 6
Row := Row + 1
return

!x::
Send, {Tab 2}{Down 3}IRL{Down}
Sleep 300
Send, {Tab}
Sleep 300
Send, {Tab 2}{Down}IRL{Down}
SLeep 300
Send, !a
return

!Numpad0::
Send, {Tab}{Enter}{Delete}!a
Sleep 300
Send, {Numpad0}
return