#Requires Autohotkey v1.1.33+
; Created by Scott Stutzman - scott-stutzman@uiowa.edu
; Version Number 2021.05.12.09.04.31

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
xcl := ComObjActive("Excel.Application") ; Variable "xcl" is now set to the last Excel work sheet that was active
InputBox, Row, What row to start on,What row would you like to start on?
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A

; !`::
F1::
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A
Clipboard := ColumnA
Send,  ^2
Sleep 500
Send, ^v
Sleep 500
Send, {Enter}
return

; !1::
F2::
Send, {Alt Down}{Down 4}{Alt Up}
Sleep 500
Send, {Tab 2}RAD IR
Sleep 500
Send, {Tab}
Sleep 500
Send, INTERVENTIONAL RADIOLOGY{Enter}
SLeep 500
Send, !a
xcl.Range("A"Row).Interior.ColorIndex := 6
Row := Row + 1
ColumnA := xcl.Range("A"Row).text
if(ColumnA = "")
{
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}
return

!x::
Send, {Tab 2}RAD IR
Sleep 300
Send, {Tab}
Sleep 300
Send, INTERVENTIONAL RADIOLOGY{Enter}
SLeep 300
Send, !a
return