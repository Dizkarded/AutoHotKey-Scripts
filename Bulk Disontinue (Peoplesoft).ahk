#Requires Autohotkey v1.1.33+
; Created by Scott Stutzman - scott-stutzman@uiowa.edu
; Version Number 2021.08.10.10.04.48

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

xcl := ComObjActive("Excel.Application") ; Variable "xcl" is now set to the most recently active Excel Workbook & Worksheet.
InputBox, Row, What row to start on,What row would you like to start on?`n`n[Hit {ESC} at anytime to exit.],,450,175
if ErrorLevel
	ExitApp
;InputBox, ZI, Enter Z (INACTIVE) reason,What is the reason for being set to Z (INACTIVE)?`n`n[Hit {ESC} at anytime to exit.],,450,175
;if ErrorLevel
;	ExitApp
ColumnA := xcl.Range("A"Row).text
ColumnH := xcl.Range("H"Row).text

F1::
xcl.Range("A"Row).Interior.ColorIndex := 6
MouseMove 11,196
Sleep 500
Click
Sleep 500
Send, ^a^c
Sleep 500
ClipBoardString := Clipboard
If InStr(ClipboardString,"Enter any information you have and click Search. Leave fields blank for a list of all values.",false)
{
	MouseMove 11,196
	Sleep 500
	Click
	Sleep 500
	Send, {Tab 8}
}

Sleep 2000
Send, %ColumnA%
Sleep 2000
MouseMove 39, 556 ; 'Searh' Button
Sleep 2000
Click
Sleep 2000
MouseMove 542, 349 ; Current status
Sleep 2000
;Send, {Down}
;Sleep 1250
;Send, {Space}
Click
Sleep 350
; MouseMove 536, 398 ; Current status / Discontinue (Home)
;MouseMove 536, 444 ; Current status / Discontinue (Work)
;Click
Send, {Down 2}{Enter}
Sleep 2000
Send, {AltDown}1{AltUp}
Sleep 2000
Send, {AltDown}2{AltUp}
Sleep 2000
xcl.Range("A"Row).Interior.ColorIndex := 4
Row := Row + 1
ColumnA := xcl.Range("A"Row).text
if(ColumnA = "")
{
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}
Sleep 1750
Send, {F1}
Clipboard := ""
Return