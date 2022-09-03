#Requires Autohotkey v1.1.09+
; Created by Scott Stutzman - scott-stutzman@uiowa.edu
; Version Number 2021.06.24.08.08.01

Z_Icon = S:\Common\AutoHotKey Scripts\Icons\Z (INACTIVE)4.ico
IfExist, %Z_Icon%
Menu, Tray, Icon, %Z_Icon%

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

xcl := ComObjActive("Excel.Application") ; Variable "xcl" is now set to the most recently active Excel Workbook & Worksheet.CoordMode, Mouse, Screen

InputBox, Row, Z (INACTIVE) comments...,What row would you like to start on?`n`n[Hit {ESC} at anytime to exit.]`n`nMake sure cursor starts in the 'SetID' Input Box,,450,200
If ErrorLevel
	ExitApp
ColumnA := xcl.Range("A"Row).text
ColumnB := xcl.Range("B"Row).text

F1::
xcl.Range("A"Row).Interior.ColorIndex := 6
WinActivate, Define Item - Google Chrome
Send, ^a^c
Sleep 500
If InStr(Clipboard,"UIOWA",true)
	Send, {Tab 2}{Delete}
Else
{	
	SoundBeep
	Reload
}	
Sleep 3000
Send, %ColumnA%
xcl.Range("A"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Sleep 3000
Send, {AltDown}1{AltUp} ; Search for part number
Sleep 3000
Send, {Tab 2} ; Status 1
Sleep 3000
xcl.Range("B"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
SendInput %ColumnB%
Send, ^a^c
Sleep 3000
If InStr(Clipboard,ColumnB)
{	
	Send, {AltDown}1{AltUp}
	Sleep 3000	
	Send, {AltDown}2{AltUp}
	Sleep 3000
}
Else
{	
	SoundBeep
	Reload
}
xcl.Range("A"Row).Interior.ColorIndex := 4
xcl.Range("B"Row).Interior.ColorIndex := 4
xcl.ActiveWindow.SmallScroll(1,0,0,0)
Row := Row + 1
ColumnA := xcl.Range("A"Row).text
ColumnH := xcl.Range("H"Row).text
If(ColumnA = "")
{
	ComObjCreate("SAPI.SpVoice").Speak("Completed")
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}
Sleep 4000
;ComObjActive("Excel.Application").ActiveWorkbook.Save()
xcl.ActiveWorkbook.Save()
Sleep 2000
Send, {F1}
Clipboard := ""
Return
/*
	Loop, 0xFF
{
	Key := Format("VK{:02X}",A_Index)
	IF GetKeyState(Key)
		Send, {%Key% Up}
}
	Reload	
*/
F2::
Send, {Tab 2}%ColumnH%
Sleep 1750
Send, {Tab 4}{Space}
Sleep 1750
Send, {Tab 4}{Space}
Sleep 1750
Send, {AltDown}1{AltUp}
Sleep 1750
Send, {AltDown}2{AltUp}
xcl.Range("A"Row).Interior.ColorIndex := 4
Row := Row + 1
ColumnA := xcl.Range("A"Row).text
ColumnH := xcl.Range("H"Row).text
if(ColumnA = "")
{
	ComObjCreate("SAPI.SpVoice").Speak("Complete")
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}
Return


F3::
Send, %ColumnH%
Sleep 2000
Send, {Tab 4}{Space}
Sleep 2000
Send, {Tab 4}{Space}
Sleep 2000
Send, {AltDown}1{AltUp}
Sleep 2000
Send, {AltDown}2{AltUp}
xcl.Range("A"Row).Interior.ColorIndex := 4
xcl.ActiveWindow.SmallScroll(1,0,0,0)
Row := Row + 1
ColumnA := xcl.Range("A"Row).text#Requires Autohotkey v1.1.09+umnH := xcl.Range("H"Row).text
if(ColumnA = "")
{
	ComObjCreate("SAPI.SpVoice").Speak("Complete")
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}
Return

*esc::
xcl := ""
ExitApp