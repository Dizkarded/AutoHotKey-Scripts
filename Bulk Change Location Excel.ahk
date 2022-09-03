#Requires Autohotkey v1.1.09+
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
InputBox, Row, What row to start on,What row would you like to start on?`n`n[Hit {ESC} at anytime to exit.]`n`nMake sure 'Correct History' is checked.,,450,200,,,,,2
If ErrorLevel
	ExitApp

ColumnA := xcl.Range("A"Row).text ; Peoplsoft Number
ColumnB := xcl.Range("B"Row).text ; Priority
ColumnC := xcl.Range("C"Row).Value ; New Location
ColumnD := xcl.Range("D"Row).text ; Supplier Number

F1::

Gosub, CGHLOC

Return

CGHLOC:
xcl.Range("A"Row).Interior.ColorIndex := 6
xcl.Range("B"Row).Interior.ColorIndex := 6
xcl.Range("C"Row).Interior.ColorIndex := 6
Send, {Tab 2}
Sleep 2000
Send, %ColumnA%
xcl.Range("A"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Sleep 2000
Send, {AltDown}1{AltUp}
Sleep 2000
Send, {AltDown}s{AltUp}
Sleep 2000
ColumnB1 := (ColumnB -1)
If (ColumnB > 0)
{
	Loop, %ColumnB1%
	{	
		Send, {AltDown}.{AltUp}
		Sleep 2000
	}	
}
xcl.Range("B"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Send, {Tab 2}
Sleep 500
Sleep 1500
Send, {Delete}%ColumnC%
xcl.Range("C"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Sleep 2000
Send, {AltDown}a{AltUp}
Sleep 2000
Send, {Enter}
Sleep 2000
Send, {AltDown}1{AltUp}
Sleep 3000
Send, {AltDown}2{AltUp}
xcl.Range("A"Row).Interior.ColorIndex := 4
xcl.Range("B"Row).Interior.ColorIndex := 4
xcl.Range("C"Row).Interior.ColorIndex := 4
Row := Row + 1
ColumnA := xcl.Range("A"Row).text
ColumnB := xcl.Range("B"Row).text
If(ColumnA = "")
{
	ComObjCreate("SAPI.SpVoice").Speak("Complete")
	MsgBox 0,Finished!, You have completed all items in list. Application will now exit.,10
	ExitApp
}
Sleep 2000
xcl.ActiveWorkbook.Save()
Sleep 1000
Gosub, CGHLOC ; Triggers this script again
Return



*esc::
xcl := ""
Colum#Requires Autohotkey v1.1.09+nA := ""
ColumnB := ""
ExitApp