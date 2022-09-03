#Requires Autohotkey v1.1.09+
; Created by Scott Stutzman - scott-stutzman@uiowa.edu
; Version Number 2021.06.24.08.08.01

Z_Icon = S:\Common\AutoHotKey Scripts\Icons\Z (INACTIVE)7.ico
IfExist, %Z_Icon%
Menu, Tray, Icon, %Z_Icon%
#Requires Autohotkey v1.1.09+
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
ColumnH := xcl.Range("H"Row).text

F1::

Gosub, BZI

Return

BZI:
Clipboard := ""
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
Sleep 2500
Send, %ColumnA%
xcl.Range("A"Row).Interior.ColorIndex := 44 ; Change backround of cell to cyan
Sleep 2500
Send, {AltDown}1{AltUp} ; Search for part number
Sleep 2500
Send, D ; Status 1
Sleep 2500
Send, {Space} ; Dismiss warning
Sleep 2500
Send, D ; Status 2
Sleep 2500
Send, ^a^c
Sleep 500
If InStr(Clipboard,"Denied Approval",True)
{	
	Send, {Tab}
	Sleep 500	
	Send, +{Tab}
	Sleep 500
	Send, D
	Clipboard :=
	Sleep 500
}
Send, {Tab}
Sleep 2500
Send, {Enter} ; Purchasing Item Attributes
Sleep 3000
Send, {Enter}
Sleep 2500
Send, {Tab 2} ; Go to Comments
Sleep 500
Send, ^a ; Highlight anything that may already be there
Sleep 500
Send, {Down}{Enter 2} ; Go to the bottom of what might be there and then go down 2 more lines
Sleep 2500
SendInput %ColumnH%
xcl.Range("A"Row).Interior.ColorIndex := 8
Sleep 2500
Send, {Tab}
Sleep 2500
Send, {Tab}{Space} ; 'Click' okay
Sleep 3500
Send, {AltDown}1{AltUp} ; 'Click' okay
Sleep 4000
Send, ^a^c
Sleep 500
If InStr(Clipboard,"Message",True)
{	
	Send, {Enter}
	Clipboard :=
	Sleep 4000
}	
Send, {AltDown}1{AltUp} ; 'Click' Save
Sleep 4000
Send, {AltDown}2{AltUp} ; 'Click' Return to Search
xcl.Range("A"Row).Interior.ColorIndex := 4
xcl.ActiveWindow.SmallScroll(1,0,0,0)
Row := Row + 1
ColumnA := xcl.Range("A"Row).text
ColumnH := xcl.Range("H"Row).text
If(ColumnA = "")
{
	ComObjCreate("SAPI.SpVoice").Speak("Complete")
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}
Sleep 4000
;ComObjActive("Excel.Application").ActiveWorkbook.Save()
xcl.ActiveWorkbook.Save()
Sleep 2500
Gosub, BZI
Return


F2::
Send, {Tab 2}%ColumnH%
Sleep 2000
Send, {Tab 4}{Space}
Sleep 2000
Send, {Tab 4}{Space}
Sleep 2000
Send, {AltDown}1{AltUp}
Sleep 2000
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
ColumnA := xcl.Range("A"Row).text
ColumnH := xcl.Range("H"Row).text
if(ColumnA = "")
{
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}
Return

*esc::
xcl := ""
Reload