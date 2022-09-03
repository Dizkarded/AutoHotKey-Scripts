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
Row := 2 ; Starts the script on the second row to avoid the header
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A


!`::
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the six digit Peoplesoft number in column A
xcl.Range("A"Row).Interior.ColorIndex := 4
Clipboard := ColumnA
Send,  ^2
Sleep 500
Send, ^v
Sleep 500
Send, {Enter}
return

!x::
Send, {Alt Down}{Down 4}{Alt Up}
Sleep 300
Send, {Alt Down}{Down 2}{Alt Up}
Sleep 300
Send, {Tab 2}RAD IR
Sleep 300
Send, {Tab}
Sleep 300
Send, INTERVENTIONAL RADIOLOGY{Enter}
SLeep 300
Send, !a
xcl.Range("A"Row).Interior.ColorIndex := 6
Row++
return