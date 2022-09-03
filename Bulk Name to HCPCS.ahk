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
	The "+" is used as a symbol for {ShIft}
	Anything after ";" is merely a comment and will not be executed in the script
*/

xcl := ComObjActive("Excel.Application") ; Variable "xcl" is now set to the last Excel work sheet that was active
InputBox, Row, Requestor/Department -> HCPCS / Starting Row,What row would you like to start on? [Hit {ESC} at anytime to exit.],,450,150,,,,,2
If ErrorLevel
	ExitApp
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to PS#
ColumnN := xcl.Range("N"Row).text ;Sets the variable "ColumnN" to Requestor / Department
ColumnH := xcl.Range("H"Row).text ;Sets the variable "ColumnH" to HCPCS Code
ColumnO := xcl.Range("O"Row).text ;Sets the variable "ColumnH" to GTIN Code

F1::
Gosub, NAMEHCPCS
Return

NAMEHCPCS:
xcl.Range("A"Row).Interior.ColorIndex := 6 ; Change backround of cell to yellow
xcl.Range("N"Row).Interior.ColorIndex := 6 ; Change backround of cell to yellow
xcl.Range("H"Row).Interior.ColorIndex := 6 ; Change backround of cell to yellow
xcl.Range("O"Row).Interior.ColorIndex := 6 ; Change backround of cell to yellow
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to PS#
ColumnN := xcl.Range("N"Row).text ;Sets the variable "ColumnN" to Requester / Department
ColumnH := xcl.Range("H"Row).text ;Sets the variable "ColumnH" to HCPCS Code
ColumnO := xcl.Range("O"Row).text ;Sets the variable "ColumnH" to GTIN Code
Send, {Tab 2}{Delete}
Sleep 500
Send, %ColumnA%
xcl.Range("A"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Sleep 2500
Send, {Enter}
Sleep 2500
Send, {Alt Down}t{Alt Up}
Sleep 2500
Send, {Tab}{Delete}
Sleep 500
Send, %ColumnO%
xcl.Range("O"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Sleep 2500
Send, {Tab 2#Requires Autohotkey v1.1.09+}{Delete}
Sleep 500
Send, %ColumnN%
xcl.Range("N"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
;Sleep 500
Send, {Tab 13}
;Sleep 500
Send, %ColumnH%
xcl.Range("H"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Send, {Alt Down}1{Alt Up}
Sleep 3000
Send, {Alt Down}2{Alt Up}
Sleep 3000
xcl.Range("A"Row).Interior.ColorIndex := 4 ; Change backround of cell to green
xcl.Range("N"Row).Interior.ColorIndex := 4 ; Change backround of cell to green
xcl.Range("H"Row).Interior.ColorIndex := 4 ; Change backround of cell to green
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
Sleep 3000
Gosub, NAMEHCPCS ; Triggers this script again

Return


*esc::
requestor=
hcpcs=

Reload
Sleep 1000 ; If successful, the reload will close this instance during the Sleep, so the line below will never be reached.
MsgBox, 4,, The script could not be reloaded. Would you like to open it for editing?
IfMsgBox, Yes, Edit
	
Loop, 0xFF
{
	Key := Format("VK{:02X}",A_Index)
	IF GetKeyState(Key)
		Send, {%Key% Up}
}

Return
