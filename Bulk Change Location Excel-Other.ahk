#Requires Autohotkey v1.1.33+
; Created by Scott Stutzman - scott-stutzman@uiowa.edu
; Version Number 2021.10.07.08.05.00

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
InputBox, Row, What row to start on,What row would you like to start on?`n`n[Hit {ESC} at anytime to exit.]`n`nMake sure 'Correct History' is checked.,,450,200
if ErrorLevel
	ExitApp

ColumnA := xcl.Range("A"Row).text
ColumnB := xcl.Range("B"Row).text

F1::
xcl.Range("A"Row).Interior.ColorIndex := 6
Send, {Tab 2}
Sleep 2500
Send, %ColumnA%
Sleep 2500
Send, {AltDown}1{AltUp}
Sleep 2500
Send, {AltDown}s{AltUp}
Sleep 2500
Loop, 13
{
	Send, {Tab}
	Sleep 500
}
Sleep 2000
Send, %ColumnB%
Sleep 2500
Send, {AltDown}a{AltUp}
Sleep 2500
Send, {Enter}
Sleep 2500
Send, {AltDown}1{AltUp}
Sleep 2500
Send, {AltDown}2{AltUp}
xcl.Range("A"Row).Interior.ColorIndex := 4
Row := Row + 1
ColumnA := xcl.Range("A"Row).text
ColumnB := xcl.Range("B"Row).text
if(ColumnA = "")
{
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}
Sleep 2500
Send, {F1}
Return



*esc::
xcl := ""
ColumnA := ""
ColumnB := ""
ExitApp