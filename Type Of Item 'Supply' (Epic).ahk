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
	The "+" is used as a symbol for {Shift}
	Anything after ";" is merely a comment and will not be executed in the script
*/
xcl := ComObjActive("Excel.Application") ; Variable "xcl" is now set to the most recently active Excel Workbook & Worksheet.
InputBox, Row, Starting Row,What row would you like to start on?`n`n[Hit {ESC} at anytime to exit.],,450,175
if ErrorLevel
	ExitApp
;InputBox, ItemType, What type of item?`n`n[Hit {ESC} at anytime to exit.],,450,300
;if ErrorLevel
;	ExitApp
ColumnA := xcl.Range("A"Row).text ; Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A
ColumnB := xcl.Range("B"Row).text ; Sets the variable "ColumnB" to the Item Type in column B

F1::
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A
xcl.Range("A"Row).Interior.ColorIndex := 6 ; Turns the backround of the cell yellow - Being worked on
; Send,  ^1 ; Reg Inventory
Send,  ^2 ; Cath Inventory
Sleep 250
Send, %ColumnA%
Sleep 250
Send, {Enter}
return

F2::
ColumnB := xcl.Range("B"Row).text ; Sets the variable "ColumnB" to the Item Type in column B
Send, {Enter}
Sleep 300
;Send, {Tab 8} ; Reg Inventory
Send, {Tab 9} ; Cath Inventory
Sleep 300
;Send, %ItemType%
Send, Supply
Sleep 300
Send, {Down}{Enter}
Sleep 300
Send, !a
xcl.Range("A"Row).Interior.ColorIndex := 4 ; Turns the backround of the cell green - Completed
Row := Row + 1
ColumnA := xcl.Range("A"Row).text
if(ColumnA = "")
{
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}
Sleep 1000
Send, {F1}
return

F3::
ColumnB := xcl.Range("B"Row).text ; Sets the variable "ColumnB" to the Item Type in column B
Send, {Enter}{Enter}
Sleep 500
Send, {Tab 8}
Sleep 500
;Send, %ItemType%
Send, %ColumnB%
Sleep 500
Send, {Down}{Enter}
Sleep 500
Send, !a
xcl.Range("A"Row).Interior.ColorIndex := 4 ; Turns the backround of the cell green - Completed
Row := Row + 1
ColumnA := xcl.Range("A"Row).text
if(ColumnA = "")
{
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}
return


^esc::
xlApp.quit
xcl := ""
ExitApp