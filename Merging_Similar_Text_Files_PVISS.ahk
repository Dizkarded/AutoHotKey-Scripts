#Requires Autohotkey v1.1.33+
;*******************************************************
; Want a clear path for learning AutoHotkey; Take a look at our AutoHotkey Udemy courses.  They're structured in a way to make learning AHK EASY
; Right now you can  get a coupon code here: https://the-Automator.com/Learn
;*******************************************************''
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance,Force
/*
xcl := ComObjActive("Excel.Application") ; Variable "xcl" is now set to the last Excel work sheet that was active
ColumnA := xcl.Range("A"Row).text ;Sets the variable "ColumnA" to the Peoplesoft number in column A
*/
SetBatchLines,-1
LineNumber := 999
CommaCount = ,
FormatTime, CurrentDateTime,,yyyyMMdd-HHmm
;********************Merge files***********************************
FileSelectFile, Files, M3, S:\Common\PeopleSoft\Data Entry\3 Data Entry To Do\LoaderQ ; M3 = Multiselect existing files. set default path to LoaderQ folder
Cancelled(Files) ;Check to see if user selected any files, if not, cancel
Out_File = %CurrentDateTime%
InputBox,New_File,New File Name,File name?,,,,,,,-1,%Out_File% ;Create file name from CurrentDateTime
Cancelled(New_File) ;If they don't give it a name, cancel
;********************Now loop over files, get paths, then combine***********************************
Loop, parse, files, `n ;parse files on new line
{
	if A_Index = 1 ;get path
		Folder_Path:=A_LoopField "\" ;Store path to folder with backslash
	Else { ;If the second one, read the whole file plus header
		If (A_Index=2)
			FileRead, Data, % *P65001  Folder_Path A_LoopField 
		Else
		{
			Loop, read, %Folder_Path%%A_LoopField% ;Loop over Additional files 
			{
				Data.= A_LoopReadLine "`r`n"	;Append each row
			}
		}
	}
}
File_Name:=InVaild_FileName_Fixer(File_Name)  ;Ensure new filename doesn't have illegal charcters
FileAppend, %Data%,%Folder_Path%%New_File%,UTF-8 ;Write file
Loop, Read, S:\Common\PeopleSoft\Data Entry\3 Data Entry To Do\LoaderQ\%New_File% ; Loop through finished file and renumber the indexing numbers to be sequential
{
	LineNumber := LineNumber +1
	LineData := A_LoopReadLine
	LineData := StrReplace(LineData, "`t")
	LineData := StrReplace(LineData, "  ", " ") ; Replace double space with single space
	LineData := StrReplace(LineData, " ,", ",") ; Replace space and comma with comma
	LineData := StrReplace(LineData, """""", """") ; Replace double-double quotes with single double quote
	LineData := StrReplace(LineData, ",""", ",") ; Replace comma double quote with comma
	LineData := StrReplace(LineData, """,Y", ",Y") ; Replace comma double quote with comma
	LineData := StrReplace(LineData, """,N", ",N") ; Replace comma double quote with comma
	LineData := StrReplace(LineData, "L"",6", "L,6") ; Replace comma double quote with comma
	FoundPos := InStr(LineData,",")
	Result := SubStr(LineData, FoundPos)
	LineData := LineNumber . Result
	StrReplace(LineData,CommaCount,CommaCount,Count)
	If Count <> 65
	{
		FoundPos2 := InStr(LineData,",",0,1,27)
		LineData := RegExReplace(LineData,CommaCount,"",,FoundPos2,FoundPos2)
	}
	/*
	MFGitemID_array := StrSplit(LineData, CommaCount)
	MFGitemID := % MFGitemID_array[17]
	Loop
	{
		If % oWorkbook.Sheets(1).Range("C"Row).Value = MFGitemID ; get value from C Column in first sheet
			MsgBox % oWorkbook.Sheets(1).Range("D"Row).Value
			Break
		Else If % oWorkbook.Sheets(1).Range("C"Row).Value = ""
			Break
	}
	*/
	FileAppend, %LineData%`r`n ,S:\Common\PeopleSoft\Data Entry\3 Data Entry To Do\LoaderQ\%Out_File%.txt ; Add each line to new file, sequentialy
	LineData := ""
}
Gosub, DeleteFile ; FileDelete, S:\Common\PeopleSoft\Data Entry\3 Data Entry To Do\LoaderQ\%Out_File% ; Delete first file
Run %Folder_Path%%Out_File%.txt ;Launch it in default text editor
Return

;********************See if the user took Action or hit cancel***********************************
Cancelled(Var){
	if (Var =""){
		MsgBox, The user pressed cancel.
		ExitApp
	}	
}

InVaild_FileName_Fixer(File_Name){
	Return File_Name:=Trim(RegExReplace(File_Name,"[\\/:*?""<>|]","_")," _") ;Get rid of invalid charachters and trim ends for underscore or space
}

DeleteFile:
FileDelete, S:\Common\PeopleSoft\Data Entry\3 Data Entry To Do\LoaderQ\%Out_File% ; Delete first file
Return