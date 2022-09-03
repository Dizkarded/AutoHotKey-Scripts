#Requires Autohotkey v1.1.09+
I_Icon = S:\Common\AutoHotKey Scripts\Icons\One Script 3.ico
IfExist, %I_Icon%
Menu, Tray, Icon, %I_Icon%

; Created by © Scott Stutzman - scott-stutzman@uiowa.edu
; Version Number 2022.02.23.10.33.26

#SingleInstance,Force ; Allows only one instance of scirpt to run at a time.

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#HotkeyModIfierTimeout 250
GroupAdd, Explore, ahk_class CabinetWClass
GroupAdd, Explore, ahk_class ExploreWClass
#Include S:\Common\AutoHotKey Scripts\Included\Email\Email.ahk
#Include S:\Common\AutoHotKey Scripts\Included\Peoplesoft\Peoplesoft.ahk
#Include S:\Common\AutoHotKey Scripts\Included\Excel\Excel.ahk
#Include S:\Common\AutoHotKey Scripts\Included\Smartsheet\Smartsheet.ahk
#Include S:\Common\AutoHotKey Scripts\Included\Date_Time\Date_Time.ahk
#Include S:\Common\AutoHotKey Scripts\Included\ScreenClip\ScreenClip.ahk
#Include S:\Common\AutoHotKey Scripts\Included\Go_Sub\Go_Sub.ahk
#Include S:\Common\AutoHotKey Scripts\Included\AutoCorrect\AutoCorrect.ahk
SetTitleMatchMode, 2

; #UseHook ;Disable when back at work!
; var2 := RegExReplace(var, "[^[:alnum:]]")
DollarSign = "$"

+esc:: ; (Shift + Escape) Reload script
Reload
Return

/*
	General information:	The "!" is used as a symbol for {Alt}
	The "#" is used as a symbol for {Windows Key}
	The "^" is used as a symbol for {Control}
	The "+" is used as a symbol for {ShIft}
	Anything after ";" is merely a comment and will not be executed in the script
*/


;__________________________________________________________________________Misc.______________________________________________________________________________________________

!/:: ; (alt + /) sends "÷"
Send, ÷
Return

;_____________________________________________________________________________________________________________________________________________________________________________

^+RButton:: ; (Control + Right Mouse Button) use after highlighting text and it will search the text in Google Search
{
	Send, ^c
	Sleep 50
	Run "https://www.google.com/search?q=%clipboard%"
}
Return

;_____________________________________________________________________________________________________________________________________________________________________________

; (Control + Space) Always on Top
^SPACE:: Winset, Alwaysontop, , A ; ctrl + space
Return

;_____________________________________________________________________________________________________________________________________________________________________________

!v:: ; (alt + v) Alternate paste of common names etc for use in creating new PS numbers

ClipboardString := Clipboard
ClipboardString := Trim(ClipboardString) ; Trim Spaces at the end of ClipboardString
/*
	If InStr(ClipboardString,"0000000300000000000")
	{
		Send, % SubStr(ClipboardString, -5)
	}
	Else If InStr(ClipboardString,"0000000000000")
	{
		Send, % LTrim(ClipboardString, "0")
	}	
	Else If InStr(ClipboardString,"000000000000")
	{
		Send, % LTrim(ClipboardString, "0")
	}
*/ ;Else
If InStr(ClipboardString,"Application Engine")
{
	PINumber := SubStr(Clipboard,1,8)
	DateTime := SubStr(Clipboard, -22)
	Send, Process Instance:{Space}%PINumber%{Space}|{Space}%DateTime%{Space}|{Space}Success{Enter}{Up}
}
Else If InStr(ClipboardString,"Process Instance:")
{
	PINumber := SubStr(Clipboard, -7)
	FormatTime, CurrentDateTime,, MM/dd/yyyy HH:mm:ss
	Send, Process Instance:{Space}%PINumber%{Space}|{Space}%CurrentDateTime%{Space}CDT{Space}|{Space}Success{Enter}{Up}
}
/*
Else If InStr(ClipboardString,"0",,7)
{
	Send, 000000%ClipboardString%
}
Else If InStr(ClipboardString,"0",,6)
{
	Send, 000000%ClipboardString%
}
*/
Else If SubStr(ClipboardString,1,7 = "0000000")
{
	ClipboardString := SubStr(ClipboardString, -5)
	Send, "0000000000000"%ClipboardString%
}
Else If SubStr(ClipboardString,1,7 = "000000")
{
	ClipboardString := SubStr(ClipboardString, -6)
	Send, "000000000000"%ClipboardString%
}
;___________Aaron Nissley___________
Else If (InStr(ClipboardString,"Nissley") OR InStr(ClipboardString,"Anesthesia"))
	Send, Aaron Nissley / Anesthesia
;___________Alicia Metcalf___________
Else If (InStr(ClipboardString,"Brunkhorst") OR InStr(ClipboardString,"Alicia") OR InStr(ClipboardString,"Metcalf"))
	Send, Alicia Metcalf / MOR
;___________Anne Grothusen___________
Else If InStr(ClipboardString,"Grothusen")
	Send, Anne Grothusen / Urology
;___________Annie Jipp___________
Else If (InStr(ClipboardString,"Jipp") OR InStr(ClipboardString,"Annie"))
	Send, Annie Jipp / MOR
;___________Ashley Byrd___________
Else If (InStr(ClipboardString,"Byrd") OR InStr(ClipboardString,"Ashley"))
	Send, Ashley Byrd / Cath Lab
;___________Becky Dill-Devor___________
Else If InStr(ClipboardString,"Dill")
	Send, R. Dill-Devor / IRL Path Lab
;___________Beth Alden___________
Else If InStr(ClipboardString,"Alden")
	Send, Beth Alden / Pathology
;___________Brandan Juhl___________
Else If (InStr(ClipboardString,"Juhl") OR InStr(ClipboardString,"bljuhl"))
	Send, Brandan Juhl / Supply Chain
;___________Cameron Daft___________
Else If InStr(ClipboardString,"Daft")
	Send, Cameron Daft/Digestive Health
;___________Christina Cherrill___________
Else If InStr(ClipboardString,"Cherrill")
	Send, Christina Cherrill / CSS
;___________Chris Maurer___________
Else If InStr(ClipboardString,"Maurer")
	Send, Chris Maurer / L & L Services
;___________Christopher Sales___________
Else If InStr(ClipboardString,"sales")
	Send, Christopher Sales / ASC
;___________Christa Venenga___________
Else If InStr(ClipboardString,"Venenga")
	Send, Christa Venenga / ASC
;___________Connie Sherman___________
Else If (InStr(ClipboardString,"Connie") OR InStr(ClipboardString,"Sherman"))
	Send, Connie Sherman / Radiology
;___________Dave Breon___________
Else If InStr(ClipboardString,"dbre_NEW")
	Send, Dave Breon / Supply Chain
;___________Dave Frederick___________
Else If InStr(ClipboardString,"Frederick")
	Send, Dave Frederick / Pathology
;___________Dave Pasbrig___________
Else If InStr(ClipboardString,"pasbrig")
	Send, Dave Pasbrig / Supply Chain
;___________Dawn Huff___________
Else If InStr(ClipboardString,"huff")
	Send, Dawn Huff / Respiratory Care
;___________Dori Pirkl___________
Else If (InStr(ClipboardString,"Pirkl") OR InStr(ClipboardString,"dori"))
	Send, Dori Pirkl / ASC
;___________Ed Phipps___________
Else If InStr(ClipboardString,"Phipps")
	Send, Ed Phipps / CSS
;___________Elizabeth LeIffert___________
Else If (InStr(ClipboardString,"Leiffert") AND InStr(ClipboardString,"Microbiology"))
	Send, E. LeIffert/Mol Microbiol Path
Else If (InStr(ClipboardString,"Leiffert") OR InStr(ClipboardString,"Elizabeth"))
	Send, Elizabeth LeIffert / Pathology
;___________James Chung___________
Else If InStr(ClipboardString,"Chung")
	Send, James Chung / ISC
;___________Jarrett Walsh___________
Else If InStr(ClipboardString,"Jarrett")
	Send, Jarrett Walsh / MOR
;___________Jeana King___________
Else If InStr(ClipboardString,"jeana")
	Send, Jeana King / SFCH Periop
;___________John Burroughs___________
Else If InStr(ClipboardString,"Burroughs")
	Send, John Burroughs / Supply Chain
;___________Joleen Lawrence___________
Else If InStr(ClipboardString,"Lawrence")
	Send, Joleen Lawrence / ASC
;___________Jonathan Hoppe___________
Else If InStr(ClipboardString,"Hoppe")
	Send, J. Hoppe / Digestive Health
;___________Jon Schlorholtz___________
Else If InStr(ClipboardString,"jon")
	Send, Jon Schlorholtz / LMC
;___________Julie Liebe___________
Else If InStr(ClipboardString,"Liebe")
	Send, Julie Liebe / MOR
;___________Jolyn Schneider___________
Else If InStr(ClipboardString,"Schneider")
	Send, Jolyn Schneider / BTC
;___________Kevin Bigbee___________
Else If InStr(ClipboardString,"bigbee")
	Send, Kevin Bigbee / PPI
;___________Kyle Harris___________
Else If InStr(ClipboardString,"kyle")
	Send, Kyle Harris / Peds Cath Lab
;___________Linda Block___________
Else If InStr(ClipboardString,"block")
	Send, Linda Block / CSS
;___________Lori Kleopfer___________
Else If InStr(ClipboardString,"Kleopfer")
	Send, Lori Kleopfer / MOR
;___________Lori Steffens___________
Else If (InStr(ClipboardString,"Steffen") OR InStr(ClipboardString,"lori/css") OR InStr(ClipboardString,"css/lori") OR InStr(ClipboardString,"Steffens - CSS"))
	Send, Lori Steffens / CSS
;___________Marcia Dragos___________
Else If (InStr(ClipboardString,"lmc / marcia") OR InStr(ClipboardString,"lmc/sub") OR InStr(ClipboardString,"marcia/lmc") OR InStr(ClipboardString,"lmc/marcia") OR InStr(ClipboardString,"marcia / LMC") OR InStr(ClipboardString,"marcia/sub") OR InStr(ClipboardString,"Marcia lmc"))
	Send, Marcia Dragos / LMC Sub
Else If (InStr(ClipboardString,"marcia") OR InStr(ClipboardString,"Dragos") OR InStr(ClipboardString,"drag_NEW"))
	Send, Marcia Dragos / Supply Chain
Else If InStr(ClipboardString,"Joleen/eye team")
	Send, Joleen / Eye Team
;___________Matt Christianson___________
Else If InStr(ClipboardString,"Christianson")
	Send, Matt Christianson/Supply Chain
;___________Melissa Vrban___________
Else If InStr(ClipboardString,"Vrban")
	Send, Melissa Vrban / ASC Nursing
;___________Mike Hoberg___________
Else If InStr(ClipboardString,"Mike/MOR")
	Send, Mike Behnami / MOR
;___________Mike Hoberg___________
Else If (InStr(ClipboardString,"Hoberg") OR InStr(ClipboardString,"hobe_NEW"))
	Send, Mike Hoberg / Supply Chain
;___________Miranda Higgins___________
Else If InStr(ClipboardString,"higgins")
	Send, Miranda Higgins / Nursing
;___________Nicolas Noiseux___________
Else If InStr(ClipboardString,"Noiseux")
	Send, Nicolas Noiseux / Orthopedics
;___________Rob Brumm___________
Else If InStr(ClipboardString,"brumm")
	Send, Rob Brumm / Radiology
;___________Ross Harrison___________
Else If (InStr(ClipboardString,"Ross") OR InStr(ClipboardString,"Harrison"))
	Send Ross Harrison / SFCH
;___________Ryan Bernemann___________
Else If (InStr(ClipboardString,"ryan/mor") OR InStr(ClipboardString,"berneman/mor"))
	Send, Ryan Bernemann / MOR
Else If InStr(ClipboardString,"ryan")
	Send, Ryan Bernemann / CSS
;___________Shelley Haganman___________
Else If (InStr(ClipboardString,"shelley") OR InStr(ClipboardString,"Haganman"))
	Send, Shelley Haganman / MOR
;___________Scott Brown___________
Else If InStr(ClipboardString,"Perfusion")
	Send, Scott Brown / Perfusion
Else If (RegExMatch(ClipboardString,"smbr_NEW")) AND (RegExMatch(ClipboardString,"Peds Cath Lab"))
	Send, Scott Brown / Peds Cath Lab
Else If (RegExMatch(ClipboardString,"smbr_NEW")) AND (RegExMatch(ClipboardString,"OB GYN"))
	Send, Scott Brown / OB-GYN
Else If (RegExMatch(ClipboardString,"smbr_NEW")) AND (RegExMatch(ClipboardString,"Cath Lab"))
	Send, Scott Brown / Cath Lab
Else If (RegExMatch(ClipboardString,"smbr_NEW")) AND (RegExMatch(ClipboardString,"Urology"))
	Send, Scott Brown / Urology
Else If (RegExMatch(ClipboardString,"smbr_NEW")) AND (RegExMatch(ClipboardString,"Research"))
	Send, Scott Brown / Research
Else If (RegExMatch(ClipboardString,"smbr_NEW")) AND (RegExMatch(ClipboardString,"DHC"))
	Send, Scott Brown / DHC
Else If (RegExMatch(ClipboardString,"smbr_NEW")) AND (RegExMatch(ClipboardString,"Respiratory Care"))
	Send, Scott Brown / Respiratory Care
Else If (RegExMatch(ClipboardString,"smbr_NEW")) AND (RegExMatch(ClipboardString,"MOR")) OR (RegExMatch(ClipboardString,"Main OR"))
	Send, Scott Brown / MOR
/*
	Else If (InStr(ClipboardString,"smbr_NEW") OR InStr(ClipboardString,"UIHC",true) OR InStr(ClipboardString,"Brown"))
		Send, Scott Brown / Supply Chain
	Else If ClipboardString contains smbr_NEW,UIHC
		Send, Scott Brown / Supply Chain
*/
;___________Scott Nibaur___________
Else If InStr(ClipboardString,"nibaur")
	Send, Scott Nibaur / CSS
;___________Steph Coleman___________
Else If InStr(ClipboardString,"Coleman")
	Send, Steph Coleman / Radiology
Else If (RegExMatch(ClipboardString,"smbr_NEW")) AND (RegExMatch(ClipboardString,"IR"))
	Send, Steph Coleman / Radiology
;___________Steve Bird___________
Else If InStr(ClipboardString,"Bird")
	Send, Steve Bird / Supply Chain
;___________Susan Hand___________
Else If InStr(ClipboardString,"Hand")
	Send, Susan Hand / Pathology
;___________Tom Gerot___________
Else If InStr(ClipboardString,"tom")
	Send, Tom Gerot / LMC
;___________Travis Suhr___________
Else If InStr(ClipboardString,"Travis")
	Send, Travis Suhr / IRL Supply Chain
;___________Wolf___________
Else If InStr(ClipboardString,"wolf")
	Send, Wolf / ASC
;___________Zac Wilson___________
Else If (InStr(ClipboardString,"Zac") OR InStr(ClipboardString,"Wilson") OR InStr(ClipboardString,"ZACH/ASC"))
	Send, Zac Wilson / ASC
;___________Alisa Buchanan___________
Else If (InStr(ClipboardString,"backorder") OR InStr(ClipboardString,"buchanana") OR InStr(ClipboardString,"replacement"))
	Send, Alisa Buchanan / BO Sub
Else If InStr(ClipboardString,"alisa")
{
	CoordMode, Mouse, Screen
	GUI, Destroy
	MouseGetPos, xpos, ypos
	xpos := xpos - 80
	ypos := ypos - 80
	GUI, Add, Radio, altsubmit gCheckAlisa vAlisaVar, Alisa Buchanan / Supply Chain
	GUI, Add, Radio, altsubmit gCheckAlisa, Alisa Buchanan / BO Sub
	GUI, Add, Radio, altsubmit gCheckAlisa, Alisa Buchanan / RC Sub
	GUI, Show, x%xpos% y%ypos%
	Return
	
	;GUIClose:
	;Return
	
	CheckAlisa:
	GUI, Submit, nohide
	If (AlisaVar = 1)
	{
		AlisaVar := "Alisa Buchanan / Supply Chain"
	}
	If (AlisaVar = 2)
	{
		AlisaVar := "Alisa Buchanan / BO Sub"
	}
	If (AlisaVar = 3)
	{
		AlisaVar := "Alisa Buchanan / RC Sub"
	}
	GUI, Destroy
	Send, %AlisaVar%
	Return
}
/*
	Else If InStr(ClipboardString,"Application Engine")
	{
		PINumber := SubStr(Clipboard,1,8)
		DateTime := SubStr(Clipboard, -22)
		Send, Process Instance:{Space}%PINumber%{Space}|{Space}%DateTime%{Space}|{Space}Success{Enter}{Up}
	}

Else If
{	
	ClipboardString := RegExReplace(ClipboardString , ",", "")
	Send, %ClipboardString%
}
*/
Else
{
	ClipClean := ClipboardString
	ClipClean2 := RegExReplace(ClipClean, "[^[:alnum:]]")
	Clipboard := ClipClean2
	Send, %ClipClean2%
}
Gosub, AllKeysUp

Return

;_____________________________________________________________________________________________________________________________________________________________________________

^+v::  

ClipClean := Clipboard
ClipClean2 := RegExReplace(ClipClean, "0000000300000000000|0000000000000|000000000000|0000000|000000|,|\$", "")

Clipboard := ClipClean2

Send, ^v

Gosub, AllKeysUp

Return

;_____________________________________________________________________________________________________________________________________________________________________________

!+v:: ; (Alt + ShIft + v) Converts Clipboard to all upper case 
{
	LowerCaseIn := Clipboard
	StringUpper UpperCaseOut, LowerCaseIn ; Clipboard manipulation code
	SendInput, %UpperCaseOut%
}

Gosub, AllKeysUp

Return

;_____________________________________________________________________________________________________________________________________________________________________________

!q:: ; (Alt + q) Sends Alt + Tab ; for working remotely
{
	Send, {Alt Down}{Tab}{Alt Up} 
}

Gosub, AllKeysUp

Return

;_____________________________________________________________________________________________________________________________________________________________________________

!+LButton:: ; (Control + ShIft + Left Mouse Button) Send, F5
{
	Send, {F5}
}

Gosub, AllKeysUp

Return

;_____________________________________________________________________________________________________________________________________________________________________________

#RButton:: ; for working remotely
Send, {Backspace}{Backspace}{Home}{Delete}

Gosub, AllKeysUp

Return

;_____________________________________________________________________________________________________________________________________________________________________________

:*:]wsr:: ; for working remotely
Send, #+{Right}

Gosub, AllKeysUp

Return

;_____________________________________________________________________________________________________________________________________________________________________________

:*:]wsl:: ; for working remotely
Send, #+{Left}

Gosub, AllKeysUp

Return

;_____________________________________________________________________________________________________________________________________________________________________________

:*:]16:: ; for working remotely
Send, {Tab 15}

Gosub, AllKeysUp

Return

;_____________________________________________________________________________________________________________________________________________________________________________

:*:]17:: ; for working remotely
Send, {Tab 17}

Gosub, AllKeysUp

Return

;_____________________________________________________________________________________________________________________________________________________________________________

:*:]mfk::
Send, 170{Tab}
Send, 70{Tab}
Send, 7890{Tab}
Send, 00000{Tab}
Send, 00000000{Tab 2}
Send, 31

Gosub, AllKeysUp

Return

;_____________________________________________________________________________________________________________________________________________________________________________

#f:: ; Radio button with choice of Family code for Peoplesoft/Excel
CoordMode, Mouse, Screen
GUI, Destroy
FormatTime, CurrentDateTime,, yyyy-MM-dd
MouseGetPos, xpos, ypos
xpos := xpos - 80
ypos := ypos - 80
GUI, Add, Radio, altsubmit gCheck vRadioGroup, &B
GUI, Add, Radio, altsubmit gCheck, &C
GUI, Add, Radio, altsubmit gCheck, &D
GUI, Add, Radio, altsubmit gCheck, &E
GUI, Add, Radio, altsubmit gCheck, &F
GUI, Add, Radio, altsubmit gCheck, &G
GUI, Add, Radio, altsubmit gCheck, &H
GUI, Add, Radio, altsubmit gCheck, &I
GUI, Add, Radio, altsubmit gCheck, &J

GUI, Show, x%xpos% y%ypos%
Return

Check:
GUI, submit, nohide
If (RadioGroup = 1)
{
	RadioGroup := "UHPS_B"
}
If (RadioGroup = 2)
{
	RadioGroup := "UHPS_C"
}
If (RadioGroup = 3)
{
	RadioGroup := "UHPS_D"
}
If (RadioGroup = 4)
{
	RadioGroup := "UHPS_E"
}
If (RadioGroup = 5)
{
	RadioGroup := "UHPS_F"
}
If (RadioGroup = 6)
{
	RadioGroup := "UHPS_G"
}
If (RadioGroup = 7)
{
	RadioGroup := "UHPS_H"
}
If (RadioGroup = 8)
{
	RadioGroup := "UHPS_I"
}
If (RadioGroup = 9)
{
	RadioGroup := "UHPS_J"
}

GUI, Destroy
Send, %RadioGroup%

Gosub, AllKeysUp

Return
;_____________________________________________________________________________________________________________________________________________________________________________

!-::
Send, {Space 5}╣╠{Left}
Sleep 300
Send, ^v

Gosub, AllKeysUp

Return

;_____________________________________________________________________________________________________________________________________________________________________________

^#!RButton::
Gosub, AllKeysUp

MsgBox 0,Finished!, All keys have been released. Click [Okay] to continue.,10
Return

;_____________________________________________________________________________________________________________________________________________________________________________

!F2::
Send, {F2}^a^c ; {End}
Sleep 500
ClipTXT := RegExReplace(Clipboard," - Copy.csv",".txt")
Sleep 500
Send, %ClipTXT%{Enter}
Sleep 500
Send, {Enter}
Sleep 500
Send, {Enter}
Return

;_____________________________________________________________________________________________________________________________________________________________________________

#/::
AppsKey & /::

ClipSLASH := Clipboard
ClipSLASH2 := RegExReplace(ClipSLASH, "/", " / ")

Clipboard := ClipSLASH2

Send, ^v

Gosub, AllKeysUp

Return

;_____________________________________________________________________________________________________________________________________________________________________________


:*:]awv::

BUOClipboard := ClipboardAll

InputBox, Paste1, Longterm Paste 1,What would you like to paste? (Single line of text only)`n`nUse the [Alt key] + x to paste.`n`n[Hit {ESC} at anytime to exit.],,650,200
If ErrorLevel
	ExitApp

InputBox, Paste2, Longterm Paste 2,What would you like to paste? (Single line of text only)`n`nUse the [Alt key] + x to paste.`n`n[Hit {ESC} at anytime to exit.],,650,200
If ErrorLevel
	ExitApp

InputBox, Paste3, Longterm Paste 3,What would you like to paste? (Single line of text only)`n`nUse the [Alt key] + x to paste.`n`n[Hit {ESC} at anytime to exit.],,650,200
If ErrorLevel
	ExitApp

!x::
CoordMode, Mouse, Screen
GUI, Destroy
MouseGetPos, xpos, ypos
xpos := xpos - 40
ypos := ypos - 40
GUI, Add, Radio, altSubmit gCheckX vPaste, %Paste1%
GUI, Add, Radio, altSubmit gCheckX, %Paste2%
GUI, Add, Radio, altSubmit gCheckX, %Paste3%
GUI, Show, x%xpos% y%ypos%
Return

CheckX:
GUI, Submit, nohide
If (Paste = 1)
{
	Clipboard := Paste1
}
If (Paste = 2)
{
	Clipboard := Paste2
}
If (Paste = 3)
{
	Clipboard := Paste3
}
GUI, Destroy
Sleep 500
Send, ^v
Return

;_____________________________________________________________________________________________________________________________________________________________________________

^+m::
Run, S:\Common\AutoHotKey Scripts\Merging_Similar_Text_FilesTest.ahk
Return

;_____________________________________________________________________________________________________________________________________________________________________________

:*:]uni:: ; Radio button with choice of Family code for Peoplesoft/Excel
CoordMode, Mouse, Screen
GUI, Destroy
FormatTime, CurrentDateTime,, yyyy-MM-dd
MouseGetPos, xpos, ypos
xpos := xpos - 80
ypos := ypos - 80
GUI, Add, Radio, altsubmit gCheckuni vRadioGroupuni, ¯\_(ツ)_/¯ (Shrug)
GUI, Add, Radio, altsubmit gCheckuni, (° ͜ʖ͡°)╭∩╮ (Flip)
GUI, Add, Radio, altsubmit gCheckuni, (ง '̀-'́)ง (Fight)
GUI, Add, Radio, altsubmit gCheckuni, ლ(-́-́ლ) (Boxer)
GUI, Add, Radio, altsubmit gCheckuni, [¬º-°]¬ (Zombi)
GUI, Add, Radio, altsubmit gCheckuni, (☞ﾟヮﾟ)☞ (You da man)
GUI, Add, Radio, altsubmit gCheckuni, (╯°□°）╯︵ ┻━┻ (Table Flip)
GUI, Add, Radio, altsubmit gCheckuni, (∩｀-')⊃━☆ﾟ.*･｡ﾟ  (PFM)
GUI, Add, Radio, altsubmit gCheckuni, 👍  (Thumbs Up)

GUI, Show, x%xpos% y%ypos%
Return

Checkuni:
GUI, submit, nohide
If (RadioGroupuni = 1)
{
	RadioGroupuni := "¯\_(ツ)_/¯"
}
If (RadioGroupuni = 2)
{
	RadioGroupuni := "(° ͜ʖ͡°)╭∩╮"
}
If (RadioGroupuni = 3)
{
	RadioGroupuni := "(ง '̀-'́)ง"
}
If (RadioGroupuni = 4)
{
	RadioGroupuni := "ლ(-́-́ლ)"
}
If (RadioGroupuni = 5)
{
	RadioGroupuni := "[¬º-°]¬"
}
If (RadioGroupuni = 6)
{
	RadioGroupuni := "(☞ﾟヮﾟ)☞"
}
If (RadioGroupuni = 7)
{
	RadioGroupuni := "(╯°□°）╯︵ ┻━┻"
}
If (RadioGroupuni = 8)
{
	RadioGroupuni := "(∩｀-')⊃━☆ﾟ.*･｡ﾟ "
}
If (RadioGroupuni = 9)
{
	RadioGroupuni := "👍"
}


GUI, Destroy
Send, %RadioGroupuni%

Loop, 0xFF
{
	Key := Format("VK{:02X}",A_Index)
	IF GetKeyState(Key)
		Send, {%Key% Up}
}

Return

;_____________________________________________________________________________________________________________________________________________________________________________

+F7::
Loop
{
	Click
	Sleep, 15000
}
Return

;_____________________________________________________________________________________________________________________________________________________________________________

^r::

Send, {Tab 17}
Sleep 500
Send, r
Sleep 500
Send, {Tab}
Return

;_____________________________________________________________________________________________________________________________________________________________________________

AppsKey & =::
Loop
{
	Click
	Sleep 10000
}
Return

;_____________________________________________________________________________________________________________________________________________________________________________

:*:](c)::©

;_____________________________________________________________________________________________________________________________________________________________________________

#IfWinActive AccessGUDID
:*:]GUD::

InputBox, CompanyName, Company Name, Please enter the first few letters of company name.`n`n[Leave blank for wildcard.]`n`n[Hit {ESC} at anytime to exit.],,450,200,,,,,
If ErrorLevel
	Return
InputBox, PartNumber, Part Number, Please enter the first part of the item ID.`n`n[Leave blank for wildcard.]`n`n[Hit {ESC} at anytime to exit.],,450,200,,,,,
If ErrorLevel
	Return
SendRaw, companyName:(
Send, %CompanyName%
SendRaw, *) AND versionModelNumber:(
Send, %PartNumber%
SendRaw, *)
Send, {Enter}
Return
#IfWinActive
;_____________________________________________________________________________________________________________________________________________________________________________