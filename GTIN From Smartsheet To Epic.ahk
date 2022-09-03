#Requires Autohotkey v1.1.09+
; Created by Scott Stutzman - scott-stutzman@uiowa.edu
; Version Number 2021.07.23.09.26.57

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
xcl := ComObjActive("Excel.Application") ; Variable "xcl" is now set to the most recently active Excel Workbook & Worksheet.
InputBox, Row, GTIN From Smartsheet To Epic,What row would you like to start on?`n`n[Hit {ESC} at anytime to exit.],,450,175,,,,,1
If ErrorLevel
	ExitApp
ColumnA := xcl.Range("A"Row).text ; Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A
ColumnJ := xcl.Range("J"Row).text ; Sets the variable "ColumnJ" to the 14 digit GTIN number in column J
ColumnC := xcl.Range("C"Row).text ; Sets the variable "ColumnC" to the description in column C

F1::
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A
ColumnC := xcl.Range("C"Row).text ; Sets the variable "ColumnC" to the description in column C
xcl.Range("A"Row).Interior.ColorIndex := 6 ; Turns the backround of the cell yellow - Being worked on
xcl.Range("J"Row).Interior.ColorIndex := 6 ; Turns the backround of the cell yellow - Being worked on
xcl.Range("C"Row).Interior.ColorIndex := 6 ; Turns the backround of the cell yellow - Being worked on
/*
If InStr(ColumnC,"cath",false)
{	
	Send, ^2
	ItemType := "Catheter"
}
Else If InStr(ColumnC,"stent",false)
{	
	Send, ^2
	ItemType := "Stent"
}
Else If InStr(ColumnC,"Occluder",false)
{	
	Send, ^2
	ItemType := "Implant"
}
Else
{
	Send, ^1
	ItemType := "Supply"
}
*/
Send, ^1
Sleep 500
Send, %ColumnA% ;^v
xcl.Range("A"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Sleep #Requires Autohotkey v1.1.09+500
Send, {Enter}
Return

F2::
Sleep 700
Send, {Enter}
ColumnJ := xcl.Range("J"Row).text ; Sets the variable "ColumnJ" to the 14 digit GTIN number in column J
/*
If InStr(ColumnC,"cath",false)
	Send, {Tab 9}
Else If InStr(ColumnC,"stent",false)
	Send, {Tab 9}
Else
	Send, {Tab 8}
Sleep 700
Send, %ItemType%
xcl.Range("C"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Sleep 700
Send, {Down}
Sleep 700
Send, {Tab 5}
Sleep 500
*/
Send, {Tab 13}
Sleep 500
Send, %ColumnJ%
xcl.Range("J"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Sleep 500
Send, !a
xcl.Range("A"Row).Interior.ColorIndex := 4 ; Turns the backround of the cell green - Completed
xcl.Range("J"Row).Interior.ColorIndex := 4 ; Turns the backround of the cell green - Completed
xcl.Range("C"Row).Interior.ColorIndex := 4 ; Turns the backround of the cell green - Completed
xcl.ActiveWindow.SmallScroll(1,0,0,0)
Row := Row + 1
ColumnA := xcl.Range("A"Row).text ; Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A
ColumnC := xcl.Range("C"Row).text ; Sets the variable "ColumnC" to the description in column C
If(ColumnA = "")
{
	ComObjCreate("SAPI.SpVoice").Speak("Complete")
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}
Else
{	
	ComObjActive("Excel.Application").ActiveWorkbook.Save()
	Sleep 1000
	Send {F1} ; Triggers this script again
}	
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
xcl.ActiveWindow.SmallScroll(1,0,0,0)
Row := Row + 1
ColumnA := xcl.Range("A"Row).text
if(ColumnA = "")
{
	ComObjCreate("SAPI.SpVoice").Speak("Complete")
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}
Return

*esc::
xlApp.quit
xcl := ""
ExitApp