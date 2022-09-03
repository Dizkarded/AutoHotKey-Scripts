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
xcl := ComObjActive("Excel.Application") ; Variable "xcl" is now set to the most recently active Excel Workbook & Worksheet.
InputBox, StartRow, Duplicate column A, What row would you like to start on?`n`n[Hit {ESC} at anytime to exit.],,450,175,,,,,1
If ErrorLevel
	Reload
Else
RowA := StartRow
RowB := RowA
ColumnA := xcl.Range("A"RowA).value ; Sets the variable "ColumnA" to the string in#Requires Autohotkey v1.1.09+ column A

F1::

Gosub, Double
Return

Double:
ColumnA := xcl.Range("A"RowA).value ;Sets the variable "ColumnA" to the string in column A
xcl.Range("B"RowB).value := ColumnA
RowB := RowB + 1
xcl.Range("B"RowB).value := ColumnA
RowB := RowB + 1
RowA := RowA + 1
ColumnA := xcl.Range("A"RowA).text ;Sets the variable "ColumnA" to the string in column A
EndRow := (RowA - StartRow)
If(ColumnA = "")
{
	RowA := RowA -1
	MsgBox 0,Finished!, %EndRow% rows have been doubled. Application will now exit.,10
	RowA := ""
	RowB := ""	
	RowB := ""	
	ExitApp
}
Else
	Loop, 0xFF
{
	Key := Format("VK{:02X}",A_Index)
	IF GetKeyState(Key)
		Send, {%Key% Up}
}
Sleep 250
Gosub, Double
Return

+F2::

Gosub, Triple
Return

Triple:
ColumnA := xcl.Range("A"RowA).text ;Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A
xcl.Range("B"RowB).value := ColumnA
RowB := RowB + 1
xcl.Range("B"RowB).value := ColumnA
RowB := RowB + 1
xcl.Range("B"RowB).value := ColumnA
RowB := RowB + 1
RowA := RowA + 1
ColumnA := xcl.Range("A"RowA).text ;Sets the variable "ColumnA" to the 18 digit Peoplesoft number in column A
EndRow := (RowA - StartRow)
If(ColumnA = "")
{
	RowA := RowA -1
	MsgBox 0,Finished!, %EndRow% rows have been tripled. Application will now exit.,10
	RowA := ""
	RowB := ""
	ExitApp
}
Else
	Loop, 0xFF
{
	Key := Format("VK{:02X}",A_Index)
	IF GetKeyState(Key)
		Send, {%Key% Up}
}
Sleep 250
Gosub, Triple
Return