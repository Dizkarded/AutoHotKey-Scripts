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
	The "+" is used as a symbol for {ShIft}
	Anything after ";" is merely a comment and will not be executed in the script
*/
xcl := ComObjActive("Excel.Application") ; Variable "xcl" is now set to the last Excel work sheet that was active
InputBox, Row, What row to start on,What row would you like to start on?`n`n[Hit {ESC} at anytime to exit.],,450,175
If ErrorLevel
	ExitApp
InputBox, Cost, What Cost,What cost would you like to change these items to?`n`n[Hit {ESC} at anytime to exit.],,450,175
If ErrorLevel
	ExitApp
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A
ColumnE := xcl.Range("E"Row).text ;Sets the variable "ColumnE" to the 18 digit Peoplesoft number in column E

; !`::
F1::
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A
ColumnE := xcl.Range("E"Row).text ;Sets the variable "ColumnE" to the 18 digit Peoplesoft number in column E
xcl.Range("A"Row).Interior.ColorIndex := 6
xcl.Range("E"Row).Interior.ColorIndex := 6
Send,  ^2
Sleep 500
Send, %ColumnA%
xcl.Range("A"Row).Interior.ColorIndex := 8
Sleep 500
Send, {Enter}

Loop, 0xFF
{
	Key := Format("VK{:02X}",A_Index)
	IF GetKeyState(Key)
		Send, {%Key% Up}
}

Return

F2::
Send, {Alt Down}{Down 2}{Alt Up}
Sleep 500
Send, {Enter}{Tab 2}
Sleep 500
Send, %ColumnE%
xcl.Range("E"Row).Interior.ColorIndex := 8
SLeep 500
Send, !a
xcl.Range("A"Row).Interior.ColorIndex := 4
xcl.Range("E"Row).Interior.ColorIndex := 4
Row := Row + 1
ColumnA := xcl.Range("A"Row).text
If(ColumnA = "")
{
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}

Loop, 0xFF
{
	Key := Format("VK{:02X}",A_Index)
	IF GetKeyState(Key)
		Send, {%Key% Up}
}

Return
