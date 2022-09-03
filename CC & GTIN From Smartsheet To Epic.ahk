#Requires Autohotkey v1.1.09+
; Created by Scott Stutzman - scott-stutzman@uiowa.edu
; Version Number 2021.07.30.08.12.17

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance, Force ; Allows only one instance of scirpt to run at a time.

BUOClipboard := Clipboard
CCiniRead := "S:\Common\AutoHotKey Scripts\CCini.ini"

xcl := ComObjActive("Excel.Application") ; Variable "xcl" is now set to the last Excel work sheet that was active
InputBox, Row, ChargeMaster -> Epic / Starting Row,What row would you like to start on?`n`n[Hit {ESC} at anytime to exit.],,450,175,,,,,1
If ErrorLevel
	ExitApp
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the Peoplesoft number in column A
ColumnE := xcl.Range("E"Row).text ;Sets the variable "ColumnE" to the MFG item ID in column E
ColumnI := xcl.Range("I"Row).text ;Sets the variable "ColumnI" to the nine digit ChargeMaster Charge Code in column C
ColumnJ := xcl.Range("J"Row).text ;Sets the variable "ColumnJ" to the fourteen digit GTIN in column D
Gosub, INI

F1::
Gosub, Start
Return

Start:
xcl.Range("A"Row).Interior.ColorIndex := 6 ; Change backround of cell in column A to yellow
xcl.Range("E"Row).Interior.ColorIndex := 6 ; Change backround of cell in column E to yellow
xcl.Range("I"Row).Interior.ColorIndex := 6 ; Change backround of cell in column I to yellow
xcl.Range("J"Row).Interior.ColorIndex := 6 ; Change backround of cell in column J to yellow
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the Peoplesoft number in column A
ColumnE := xcl.Range("E"Row).text ;Sets the variable "ColumnE" to the MFG item ID number in column E
ColumnI := xcl.Range("I"Row).text ;Sets the variable "ColumnI" to the nine digit ChargeMaster Charge Code in column C
ColumnJ := xcl.Range("J"Row).text ;Sets the variable "ColumnJ" to the fourteen digit GTIN in column D

If (ColumnI ~= GSRegInv) ; "027000001|027200000|027200026|027200027|027200029|027200041|027200049|027200072|027500001|027500002|027500003|027500011|027600001|027600007|027600010|027800000|027800002|027800004|027800009|027800011|027800013|027800014|027800016|027800018|027800019|27800020|027800050|027800051|027800053|027800056|027800058|027800071|027800074|027800084|027800091|027810543|027810544|062400075|062400081|062400083")
	;If Column H is part of Regular Inventory
{
	If (WinActive("Hyperspace"))
	{
		Send,  ^1
		Sleep 500
		Send, %ColumnA% ; Peoplesoft number in column A
		xcl.Range("A"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
		xcl.Range("E"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan		
		Sleep 500
		Send, {Enter}
	}
	
	Loop, 0xFF
	{
		Key := Format("VK{:02X}",A_Index)
		IF GetKeyState(Key)
			Send, {%Key% Up}
	}
}
Else
	;If Column H is part of Cath Inventory
{
	If (WinActive("Hyperspace"))
	{
		Send,  ^2
		Sleep 500
		Send, %ColumnA% ; Peoplesoft number in column A
		xcl.Range("A"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
		xcl.Range("E"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan		
		Sleep 500
		Send, {Enter}
	}
	
	Loop, 0xFF
	{
		Key := Format("VK{:02X}",A_Index)
		IF GetKeyState(Key)
			Send, {%Key% Up}
	}
}


Gosub, PGC
Return


PGC:
Sleep 2000
PixelGetColor, ColorVar, 332, 86, RGB

If (ColorVar = 0x3399FF)
{
	Gosub, Singular
}	
Else
{
	Gosub, Multiple
}	

Return


Singular:

;F2::
Send, {Enter}
Sleep 300
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the Peoplesoft number in column A
ColumnI := xcl.Range("I"Row).text ;Sets the variable "ColumnI" to the nine digit ChargeMaster Charge Code in column H
;The ColumnI2 variable will be entered into the "Type of item:" field
If (ColumnI ~= GSSupply)
{
	ColumnI2 := "Supply" ;Sets the variable "ColumnI2" to "Supply"
}
Else If (ColumnI ~= GSCatheter)
{
	ColumnI2 := "Catheter" ;Sets the variable "ColumnI2" to "Catheter"
}
Else If (ColumnI ~= GSBalloon)
{
	ColumnI2 := "Balloon" ;Sets the variable "ColumnI2" to "Balloon"
}
Else If (ColumnI ~= GSGuidewire)
{
	ColumnI2 := "Guidewire" ;Sets the variable "ColumnI2" to "Guidewire"
}
Else If (ColumnI ~= GSSheath)
{
	ColumnI2 := "Sheath" ;Sets the variable "ColumnI2" to "Sheath"
}
Else If (ColumnI ~= GSStentDES)
{
	ColumnI2 := "Stent DES" ;Sets the variable "ColumnI2" to "Stent DES"
}
Else If (ColumnI ~= GSStentBMS)
{
	ColumnI2 := "Stent BMS" ;Sets the variable "ColumnI2" to "Stent BMS"
}
Else If (ColumnI ~= GSStent)
{
	ColumnI2 := "Stent" ;Sets the variable "ColumnI2" to "Stent"
}
Else If (ColumnI ~= GSImplant)
{
	ColumnI2 := "Implant" ;Sets the variable "ColumnI2" to "Implant"
}
Else If (ColumnI ~= GSWatchman) ; "Watchman"
{
	ColumnI2 := "Implant" ;Sets the variable "ColumnI2" to "Implant"
}
Else If (ColumnI ~= GSMitraclip) ; "Mitraclip"
{
	ColumnI2 := "Implant" ;Sets the variable "ColumnI2" to "Implant"
}
Else If (ColumnI ~= GSTissue) ; "Tissue"
{
	ColumnI2 := "Implant" ;Sets the variable "ColumnI2" to "Implant"
}
Else If (ColumnI ~= GSDressing) ; "Dressing"
{
	ColumnI2 := "Dressing" ;Sets the variable "ColumnI2" to "Implant"
}
Else ; If none of the above "If" statements work then a message box is displayed stating the Charge Code that needs to be added to this script, adds charge code to Clipboard and then exits script
{
	MsgBox, 4096,New Item Type Needed!, Let Scott know to add:`n %ColumnI% to the INI file.,10
	Clipboard := %ColumnI%|
	ExitApp
}

If ColumnI2 ~= GSCathInv ; Cath inventory
{
	Send, {Tab 9}
}
Else ; Regular inventory
{
	Send, {Tab 8}
}
Sleep 500
Send, %ColumnI2%
If (ColumnI = "027810543")
{
	Send, {Tab}Watchman
}
Else If (ColumnI = "027810544")
{
	Send, {Tab}Mitraclip
}
Else If (ColumnI = "027800005")
{
	Send, {Tab}Tissue
}
Send, {Down}
Sleep 500
Send, {Tab 5}%ColumnJ% ; Fourteen digit GTIN in column J
xcl.Range("J"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Sleep 500
Send, {Tab 4}{Alt Down}{Down}{Down}{Alt Up}
Sleep 500
Send, {Space}{Tab 3}
Sleep 500
Send, %ColumnI% ; Nine Digit ChargeMaster Charge Code in column I
xcl.Range("I"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Sleep 500
Send, {Tab}{Tab 4}{Space}
Sleep 1000
Send, {Alt Down}A{Alt Up}
Sleep 500
xcl.Range("A"Row).Interior.ColorIndex := 4 ; Change backround of cell to green
xcl.Range("E"Row).Interior.ColorIndex := 4 ; Change backround of cell to green
xcl.Range("I"Row).Interior.ColorIndex := 4 ; Change backround of cell to green
xcl.Range("J"Row).Interior.ColorIndex := 4 ; Change backround of cell to green
xcl.ActiveWindow.SmallScroll(1,0,0,0)
Row := Row + 1 ; Set up for next Peoplesoft number
ColumnA := xcl.Range("A"Row).text 
Send, {Control Up}{Alt Up}
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
Sleep 2000
Gosub, Start ; Triggers this script again
Return

Multiple:

; F3:: ; If there was more than one item to choose from in the 'F1' area and the correct item is already highlighted
Send, {Enter}
Sleep 300
Send, {Enter}
Sleep 300
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the six digit Peoplesoft number in column A
ColumnI := xcl.Range("I"Row).text ;Sets the variable "ColumnI" to the nine digit ChargeMaster Charge Code in column H
;The ColumnI2 variable will be pasted into the "Type of item:" field
If (ColumnI ~= GSSupply)
{
	ColumnI2 := "Supply" ;Sets the variable "ColumnI2" to "Supply"
}
Else If (ColumnI ~= GSCatheter)
{
	ColumnI2 := "Catheter" ;Sets the variable "ColumnI2" to "Catheter"
}
Else If (ColumnI ~= GSBalloon)
{
	ColumnI2 := "Balloon" ;Sets the variable "ColumnI2" to "Balloon"
}
Else If (ColumnI ~= GSGuidewire)
{
	ColumnI2 := "Guidewire" ;Sets the variable "ColumnI2" to "Guidewire"
}
Else If (ColumnI ~= GSSheath)
{
	ColumnI2 := "Sheath" ;Sets the variable "ColumnI2" to "Sheath"
}
Else If (ColumnI ~= GSStentDES)
{
	ColumnI2 := "Stent DES" ;Sets the variable "ColumnI2" to "Stent DES"
}
Else If (ColumnI ~= GSStentBMS)
{
	ColumnI2 := "Stent BMS" ;Sets the variable "ColumnI2" to "Stent BMS"
}
Else If (ColumnI ~= GSStent)
{
	ColumnI2 := "Stent" ;Sets the variable "ColumnI2" to "Stent"
}
Else If (ColumnI ~= GSImplant)
{
	ColumnI2 := "Implant" ;Sets the variable "ColumnI2" to "Implant"
}
Else If (ColumnI ~= GSWatchman) ; "Watchman"
{
	ColumnI2 := "Implant" ;Sets the variable "ColumnI2" to "Implant"
}
Else If (ColumnI ~= GSMitraclip) ; "Mitraclip"
{
	ColumnI2 := "Implant" ;Sets the variable "ColumnI2" to "Implant"
}
Else If (ColumnI ~= GSTissue) ; "Tissue"
{
	ColumnI2 := "Implant" ;Sets the variable "ColumnI2" to "Implant"
}
Else If (ColumnI ~= GSDressing) ; "Dressing"
{
	ColumnI2 := "Dressing" ;Sets the variable "ColumnI2" to "Dressing"
}
Else ; If none of the above "If" statements work then a message box is displayed stating the Charge Code that needs to be added to this script, adds charge code to Clipboard and then exits script
{
	MsgBox, 4096,New Item Type Needed!, Let Scott know to add:`n %ColumnI% to the INI file.,30
	Clipboard := %ColumnI%|
	ExitApp
}

If ColumnI2 ~= GSCathInv ; Cath inventory
{
	Send, {Tab 9}
}
Else ; Regular inventory
{
	Send, {Tab 8}
}
Sleep 500
Send, %ColumnI2%
If (ColumnI = "027810543")
{
	Send, {Tab}Watchman
}
Else If (ColumnI = "027810544")
{
	Send, {Tab}Mitraclip
}
Else If (ColumnI = "027800005")
{
	Send, {Tab}Tissue
}
Send, {Down}
Sleep 500
Send, {Tab 5}%ColumnJ% ; Fourteen digit GTIN from column J
xcl.Range("J"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Sleep 500
Send, {Tab 4}{Alt Down}{Down}{Down}{Alt Up}
Sleep 500
Send, {Space}{Tab 3}
Sleep 500
Send, %ColumnI% ; Nine Digit ChargeMaster Charge Code
xcl.Range("I"Row).Interior.ColorIndex := 8 ; Change backround of cell to cyan
Sleep 500
Send, {Tab}{Tab 4}{Space}
Sleep 1000
Send, {Alt Down}A{Alt Up}
Sleep 500
xcl.Range("A"Row).Interior.ColorIndex := 4 ; Change backround of cell to green
xcl.Range("E"Row).Interior.ColorIndex := 4 ; Change backround of cell to green
xcl.Range("I"Row).Interior.ColorIndex := 4 ; Change backround of cell to green
xcl.Range("J"Row).Interior.ColorIndex := 4 ; Change backround of cell to green
xcl.ActiveWindow.SmallScroll(1,0,0,0)
Row := Row + 1 ; Set up for next Peoplesoft number
ColumnA := xcl.Range("A"Row).text 
Send, {Control Up}{Alt Up}
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
Sleep 2000
Gosub, Start
Return

*esc::
xcl := ""
Clipboard := BUOClipboard

Loop, 0xFF
{
	Key := Format("VK{:02X}",A_Index)
	IF GetKeyState(Key)
		Send, {%Key% Up}
}

Reload

INI:

IniRead, GSRegInv, %CCiniRead%, RegInv, RegInv,
IniRead, GSCathInv, %CCiniRead%, CathInv, CathInv,
IniRead, GSSupply, %CCiniRead%, Supply, Supply,
IniRead, GSCatheter, %CCiniRead%, Catheter, Catheter,
IniRead, GSBalloon, %CCiniRead%, Balloon, Balloon,
IniRead, GSGuidewire, %CCiniRead%, Guidewire, Guidewire,
IniRead, GSSheath, %CCiniRead%, Sheath, Sheath,
IniRead, GSStentDES, %CCiniRead%, StentDES, StentDES,
IniRead, GSStentBMS, %CCiniRead%, StentBMS, StentBMS,
IniRead, GSStent, %CCiniRead%, Stent, Stent,
IniRead, GSImplant, %CCiniRead%, Implant, Implant,
IniRead, GSWatchman, %CCiniRead%, Watchman, Watchman,
IniRead, GSMitraclip, %CCiniRead%, Mitraclip, Mitraclip,
IniRead, GSTissue, %CCiniRead%, Tissue, Tissue,
IniRead, GSDressing, %CCiniRead%, Dressing, Dressing,

Return