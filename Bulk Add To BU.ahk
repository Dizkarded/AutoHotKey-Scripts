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

xcl := ComObjActive("Excel.Application") ; Variable "xcl" is now set to the last Excel work sheet that was active
InputBox, Row, Bulk Add To BU / Starting Row,What row would you like to start on?`n`n[Hit {ESC} at anytime to exit.],,450,175,,,,,1
If ErrorLevel
	ExitApp
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the Peoplesoft number in column A
ColumnB := xcl.Range("B"Row).text ;Sets the variable "ColumnE" to the MFG item ID in column E

F1::

Gosub, BABU

Return

BABU:
xcl.Range("A"Row).Interior.ColorIndex := 6 ; Change backround of cell in column A to yellow
xcl.Range("B"Row).Interior.ColorIndex := 6 ; Change backround of cell in column E to yellow
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the Peoplesoft number in column A
ColumnB := xcl.Range("B"Row).text ;Sets the variable "ColumnE" to the MFG item ID number in column E
Send, {ShiftDown}{End}{Delete}
Sleep 2000
Send, %ColumnA%{Tab}
xcl.Range("A"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Sleep 2000
Send, %ColumnB%
xcl.Range("B"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Sleep 2000
Send, {AltDown}1{AltUp}
Sleep 3000
Send, {Tab}100
Sleep 2000
Send, {AltDown}i{AltUp}
Sleep 2000
Send, {Tab}{Enter}
Sleep 2000
Send, {Tab 2}
Sleep 2000
Send, ALL
Sleep 2000
Send, {AltDown}1{AltUp}
;xcl.Range("A"Row).Interior.ColorIndex := 4 ; Change backround of cell to green
;xcl.Range("B"Row).Interior.ColorIndex := 4 ; Change backround of cell to green
xcl.ActiveWindow.SmallScroll(1,0,0,0)
Row := Row + 1
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to PS#
If(ColumnA = "")
{
	ComObjCreate("SAPI.SpVoice").Speak("Complete")
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp	
}
Else
	
Loop, 0xFF
{
	Key := Format("VK{:02X}",A_Index)
	IF GetKeyState(Key)
		Send, {%Key% Up}
}

ComObjActive("Excel.Application").ActiveWorkbook.Save()
Sleep 2000
MouseMove 644, 637
Sleep 1000
Click
Sleep 1000
RowU = Row - 1
xcl.Range("A"RowU).Interior.ColorIndex := 4 ; Change backround of cell to green
xcl.Range("B"RowU).Interior.ColorIndex := 4 ; Change backroun#Requires Autohotkey v1.1.09+d of cell to green
Gosub, BABU ; Triggers this script again

;Return
Return


*esc::
Gosub, AllKeysUp21

ExitApp


AllKeysUp21:

Loop, 0xFF
{
	Key := Format("VK{:02X}",A_Index)
	IF GetKeyState(Key)
		Send, {%Key% Up}
}
Return