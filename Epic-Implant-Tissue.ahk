#Requires Autohotkey v1.1.09+
; Created by Scott Stutzman - scott-stutzman@uiowa.edu
; Version Number 2021.11.08.10.58.03

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
xcl := ComObjActive("Excel.Application") ; Variable "xcl" is now set to the most recently active Excel Workbook & Worksheet.
InputBox, Row, Epic-Implant-Tissue,What row would you like to start on?`n`n[Hit {ESC} at anytime to exit.],,450,175,,,,,2
If ErrorLevel
	ExitApp
ColumnA := xcl.Range("A"Row).text ; Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A

F1::
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A
xcl.Range("A"Row).Interior.ColorIndex := 6 ; Turns the backround of the cell yellow - Being worked on
Send,  ^1
Sleep 500
Send, %ColumnA%
xcl.Range("A"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Sleep 500
Send, {Enter}
Return

F2::
Send, {Enter}
Send, {Tab 8}
Sleep 500
Send, Implant{Tab}Tissue
Sleep 500
Send, {Alt Down}
Sleep 500
Send, {Down 4}
Sleep 500
Send, {Alt Up}
Sleep 500
Send, {Tab}{Space}
Sleep 500
Send, ASC{Down}
Sleep 500

Gosub, Accept

Send, !d
Sleep 500

Gosub, Accept

Send, y{tab}
Sleep 500

Gosub, Accept

Send, {Tab}
Sleep 500
Send, {Space}
Sleep 500
Send, SFCH{Down}
Sleep 500

Gosub, Accept

Send, !d
Sleep 500

Gosub, Accept

Send, y{tab}
Sleep 500

Gosub, Accept

Gosub, Accept

xcl.Range("A"Row).Interior.ColorIndex := 4 ; Turns the backround of the cell green - Completed
Row := Row + 1
ColumnA := xcl.Range("A"Row).text
If(ColumnA = "")
{
	ComObjCreate("#Requires Autohotkey v1.1.09+SAPI.SpVoice").Speak("Complete")
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}
Else
	Send {F1} ; Triggers this script again
Return

F3::
Send, {Enter}
Sleep 300
Send, {Enter}
ColumnJ := xcl.Range("J"Row).text ; Sets the variable "ColumnJ" to the 14 digit GTIN number in column I
Send, {Tab 13}
Sleep 300
Send, {Down}%ColumnJ%
Sleep 300
Send, !a
xcl.Range("A"Row).Interior.ColorIndex := 4 ; Turns the backround of the cell green - Completed
xcl.Range("J"Row).Interior.ColorIndex := 4 ; Turns the backround of the cell green - Completed
Row := Row + 1
ColumnA := xcl.Range("A"Row).text
If(ColumnA = "")
{
	ComObjCreate("SAPI.SpVoice").Speak("Complete")
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}
Return

Accept:
Send, {Alt Down}
Sleep 500
Send, a
Sleep 500
Send, {Alt Up}
Sleep 500
Return

*esc::
xlApp.quit
xcl := ""
Exit

