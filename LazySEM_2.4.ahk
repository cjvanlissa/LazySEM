;	 Copyright 2014-2017 Caspar van Lissa
; 	 LazySEM is free software: you can redistribute it and/or modify
;    it under the terms of the GNU General Public License as published by
;    the Free Software Foundation, either version 3 of the License, or
;    (at your option) any later version.
;
;    LazySEM is distributed in the hope that it will be useful,
;    but WITHOUT ANY WARRANTY; without even the implied warranty of
;    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;    GNU General Public License for more details.
;
;    You should have received a copy of the GNU General Public License
;    along with this program.  If not, see <http://www.gnu.org/licenses/>.

global OldClip:= ""

;
;Store/restore old clipboard
;
StoreClip(){
	OldClip := ClipBoard
	Clipboard=
	return
}
RestoreClip(){
	ClipBoard := OldClip
	return
}

getSelection(){
	ClipBoard=
	sleep, 50
	Sendinput {Ctrl down}c{ctrl up}
     	ClipWait, 2
	return
}


;
;Function to sort an array by value
;
SortArrayByValue(array,options="",delimiter="`n"){
   If InStr(Options,"ByKey"){
      StringReplace,options,options,ByKey
      for k in array
         out.= k delimiter
      StringTrimRight,out,out,1
      Sort,out,%options%
      Loop,Parse,out,%delimiter%
         out_.=Array[A_LoopField] Delimiter
      out:=out_
   } else {
      for k, v in array
         out.= v delimiter
      StringTrimRight,out,out,1
      Sort,out,%options%
   }
   return out
}

;
;Function to check if value is in Object
;

hasValue(haystack, needle) {
    if(!isObject(haystack))
        return false
    if(haystack.Length()==0)
        return false
    for k,v in haystack
        if(v==needle)
            return true
    return false
}

;
;Function to convert array to string
;
arrayToString(theArray)
{	string := "{"
	for key, value in theArray
	{	if(A_index != 1)
		{	string .= ","
		}
        if key is number
        {   string .= key ":"
        } else if(IsObject(key))
        {   string .= arrayToString(key) ":"
        } else
        {   key := escapeSpecialChars(key)
            string .=  """" key """:" 
        }
        if value is number
        {   string .= value
        } else if (IsObject(value))
		{	string .= arrayToString(value)
		} else
		{	value := escapeSpecialChars(value)
			string .=  """" value """"
		}
	}
	return string "}"
}
escapeSpecialChars(theString, reverse := false)
{	unEscaped := ["""", "``", "`r", "`n", ",", "%", ";", "::", "`b", "`t", "`v", "`a", "`f"]
	escaped := ["""""", "````", "``r", "``n", "``,", "``%", "``;", "``::", "``b", "``t", "``v", "``a", "``f"]
 
    search := reverse ? escaped : unEscaped
    replace := reverse ? unEscaped : escaped
 
	for index, s in search
	{	StringReplace, theString, theString, % s, % replace[index], All
	}
	return theString
}

;
;Function to convert string to array
;
stringToArray(theString)
{	
	if(theString == "{}")
	{
		return {}
	}
        if(RegExMatch(theString, "\R") || instr(theString, "{") != 1 || instr(theString, "}", true, 0) != strlen(theString))
	{ 	return false
	}
    returnArray := object()
    start := 2
    Loop
    {   valueString := getNextValue(theString, start) 
        if(valueString == false)
        {   ;invalid value for key
            break
        }
        key := valueString[1]
        start := valueString[2]
        if(RegExMatch(theString, "\s*:", "", start) != start++)
        {   ;no ':' after key
            break
        }
        valueString := getNextValue(theString, start)
         if(valueString == false)
        {   ;invalid value for value
            break
        }
        value := valueString[1]
        start := valueString[2] 
        returnArray.insert(key, value)
        if(RegExMatch(theString, "\s*}", "", start) == start)
        {   ;closing brace indiacates end of the object
            return returnArray
        } else
        {   start := InStr(theString, ",", true, start)
            if(start == 0)
            {   ;no closing brace or comma before the next var
                break
            }
            start++
        }
    }
    return false
}
getNextValue(ByRef string, start)
{   if(RegExMatch(string, "\s*[+-]{0,1}[\d\.]", "", start) == start)
    {   ;it's a number
        start := RegExMatch(string, "[+-]{0,1}[\d\.]", "", start)
        end := regexmatch(string, "[^\d\.+-]", value, start)
        return [substr(string, start, end - start), end]
    }
    if(RegExMatch(string, "\s*""", "", start) == start)
    {   ;it is a string
        check := start := RegExMatch(string, """", "", start) + 1
        Loop
        {   ;find the next "
            end := InStr(string, """", true, check)
            ;check if the " found is actually an escaped " (ie "")
            if(end == instr(string, """""", true, check))
            {   ;indicates an escaped "
                check := end + 2
            } else
            {   break
            }
        }
        return [escapeSpecialChars(substr(string, start, end - start), reverse := true), end + 1]
    }
    if(RegExMatch(string, "\s*\{", "", start) == start)
    {   
        ;it is another object!
        if(RegExMatch(string, "}", "", start) == start +1)
        {
            ;the other object is an emtpy object
            return [{}, start + 2]
        }
        start := instr(string, "{", true, start)
        end := start + 1
        ;if we find an { then we need to find an additional }
        braceCount := 0
        ;braces within "'s are ignored
        ignoreBraces := false
        ;find the closing brace
        while(end := RegExMatch(string, "[""\{}]", found, end))
        {   if(found == """")
            {   ignoreBraces := ! ignoreBraces
            } else if(found == "{")
            {   braceCount++
            } else if(found == "}" && ignoreBraces == false)
            {   if(braceCount == 0)
                {   break
                }
                braceCount--
            }
            end++
        }
        if(end == 0)
        {   MsgBox end is 0
            return false
        }
        ;MsgBox % substr(string, start, end - start + 1)
        value := stringToArray(substr(string, start, end - start + 1))
        if(value)
        {   return [value , end + 1 ]
        }
    }
    MsgBox didn't start with a good value
    return false
}

^`::
{
Send {VOLUME_MUTE}
return
}

^+x::
{
	FormatTime, CurrentDateTime,, dd-MM-yy
	SendInput %CurrentDateTime%_
	return
}

^+h::
{
;
;Highlight significant p-values in Excel sheet 
;
SetTitleMatchMode, 2
IfWinNotActive, Excel
	{
	MsgBox Excel must be open to highlight signicant values
	return
	}
Clipboard =
send {control down}c{control up}
ClipWait, 2

IfNotInString, Clipboard, `r`n
	{
	MsgBox Several rows must be selected to highlight significant values
	}

StringReplace, Clipboard, Clipboard, `r`n, `r`n, UseErrorLevel
count := ErrorLevel
send {down}{up}

Loop, %count%
	{
	Clipboard =
	send {control down}c{control up}
	Clipwait, 2
	IfInString, Clipboard, p =
		{
		pvalue := RegExReplace(Clipboard, ".*p = \d?(\.\d{1,2}).*", "$1")
		
		if(pvalue<.05)
			{
			
			send {control down}b{control up}
			
			}
		}
	send {down}
	}

return
}


^+g::
{

Clipboard =
send {control down}c{control up}
IfInString, Clipboard, );
	{
	ModelText := Clipboard
	EditText := Clipboard
	}
else
	{
	IfInString, ModelText, );
		EditText := ModelText
	else
		{
		MsgBox No valid model syntax selected or in memory!
		return
		}
	}
InputBox, Ext, Group extension?
StringReplace, EditText, EditText, MODEL:, MODEL %Ext%:, all
StringReplace, EditText, EditText, );, %Ext%);, all
Clipboard := EditText
return
}


^+3::
{
;
;Wrap in hashtags
;
StoreClip()

Clipboard=
SetTitleMatchMode 2
IfWinActive, Mplus
	CommentChar:="`!"
IfWinActive, TeXstudio
	CommentChar:="`%"
IfWinActive, RStudio
	CommentChar:="`#"

getSelection()
Outstring=
Hashtags=
NumHashtags := StrLen(Clipboard)+4
if Clipboard=
	{
	RestoreClip()
	Return
	}

Loop, %NumHashtags%
	{ 
		Hashtags := Hashtags . "`#" 
	} 

Outstring := Hashtags . "`r`n" . "`# " . Clipboard . " `#`r`n" . Hashtags . "`r`n"
Clipboard := Outstring
Sendinput {Ctrl down}v{ctrl up}
sleep, 20
RestoreClip()
return
}




^+1::
{
;
;Comment selection in MPlus
;
StoreClip()

Clipboard=
SetTitleMatchMode 2
IfWinActive, Mplus
	CommentChar:="`!"
IfWinActive, TeXstudio
	CommentChar:="`%"
IfWinActive, RStudio
	CommentChar:="`#"

Outstring =
getSelection()

if Clipboard=
	{
	send %CommentChar%
	RestoreClip()
	Return
	}
else
	{
	IfInString, Clipboard, %CommentChar%
		StringReplace, Outstring, Clipboard, %CommentChar%,, All
	else
		{
		;Loop, parse, Clipboard, "`r`n"
		;	{
		;	if (A_Loopfield="")
		;		continue
		;	Outstring:=Outstring . CommentChar . A_Loopfield . "`r`n"
		;	}
		;MsgBox % Clipboard
		;Outstring := RegExReplace(Clipboard, "`aim)^(.*)$", "<begining of line>$1<end of line>`n")
		;RegExReplace(Clipboard, "^", CommentChar, OutputVarCount =numreps)
		;MsgBox % numreps
		Outstring=
		Loop, parse, clipboard, `n, `r
			{
				if (A_Loopfield="")
					continue
				Outstring=%Outstring%%CommentChar%%A_LoopField%`n
			}
		}
	Clipboard := Outstring
	Sendinput {Ctrl down}v{ctrl up}
	sleep, 20
	RestoreClip()
	return
}
}


^+d::
{
;
;Remove duplicates from selected text
;
Clipboard=
Outstring =
labels := Object()
send ^c
clipwait, 2
Loop, parse, Clipboard, "`r`n"
	{
	if (A_Loopfield="")
		continue
	if !hasValue(labels, A_Loopfield) {
		labels.Push(A_LoopField)
		}
	}
numreps:=labels.MaxIndex()
Loop, %numreps%
	{
	Outstring:=Outstring . labels[A_Index] . "`r`n"
	}
Clipboard := Outstring
return
}


 

^+s::
{
Clipboard=
Outstring =
labels := Object()
send ^c
clipwait, 2
Loop, parse, Clipboard, "`r`n"
	{
	if (A_Loopfield="")
		continue
	labels.Insert(A_LoopField)
	}
Clipboard := SortArrayByValue(labels)
Sendinput {Ctrl down}v{ctrl up}
return
}


^+c::
{
;
;Turn a list of variables (space-delimited) into a series of WITH statements to get all correlations between these variables
;
SetTitleMatchMode 2
Clipboard=
Outstring=
labels := Object()
send ^c
clipwait, 2

StringSplit, labels, Clipboard, %A_Space%

Loop, %labels0%
	{
    curit:=A_Index
	numreps:=labels0-A_Index
	Loop, %numreps%
		{
		withit :=curit+A_Index
		Outstring:=Outstring . labels%curit% . " WITH " . labels%withit% . ";`r`n"
		}
	}
Clipboard := Outstring
Sendinput {Ctrl down}v{ctrl up}
return
}



^+-::
{
;
;Expand or abbreviate MPlus variable names
;
;Initialize variables
StoreClip()
Outstring :=
Input :=
Clipboard=
send ^c
ClipWait, 2

Input = %Clipboard%
CurVar :=
Num :=0
FirstVar:=
FirstNum:=
;Clean up input
StringReplace, Input, Input, `r`n, %A_Space% , All
StringReplace Input, Input, %A_Space%%A_Space%, %A_Space%, All
StringReplace Input, Input, %A_Space%%A_Space%, %A_Space%, All
StringReplace Input, Input, %A_Space%%A_Space%, %A_Space%, All
StringReplace Input, Input, %A_Space%%A_Space%, %A_Space%, All
StringReplace Input, Input, %A_Space%%A_Space%, %A_Space%, All

IfInString, Input, -
	{
    Loop, parse, Input, %A_Space%
		{
		if (InStr(A_Loopfield, "-"))
			{
			FirstVar:= RegExReplace(SubStr(A_Loopfield, 1, InStr(A_Loopfield, "-") - 1), "[0-9]+$", "")
			RegexMatch(SubStr(A_Loopfield, 1, InStr(A_Loopfield, "-") - 1), "([0-9]+$)", FirstNum)
			RegexMatch(A_Loopfield, "([0-9]+$)", Num)
			Num := Num-FirstNum+1
			Loop, %Num%
				{
				Outstring := Outstring . FirstVar . FirstNum+A_Index-1 . A_Space
				}
			}
		else
			Outstring := Outstring . A_Loopfield . A_Space
		}
	Clipboard := Outstring
	Sendinput {Ctrl down}v{ctrl up}
	return
	}
else
	{
	;Process input
	Loop, parse, Input, %A_Space%
		{
		if (RegexMatch(A_Loopfield, "^[0-9a-zA-Z_]+[0-9]+$")>0)				;If it's a variable with a number at the end
			{
			if Num = 0
				{
				FirstVar:= RegExReplace(A_Loopfield, "[0-9]+$", "")
				RegexMatch(A_Loopfield, "([0-9]+$)", FirstNum)
				Num :=1
				continue
				}
			CurVar:=RegExReplace(A_Loopfield, "[0-9]+$", "")		
			if (CurVar = FirstVar)										;If curvar is the same as firstvar, increment num and continue loop
				{
				Num := Num+1
				}
			else
				{
				Outstring := Outstring . FirstVar . FirstNum . "-" . FirstVar . FirstNum+Num-1 . A_Space
				FirstVar:= CurVar
				RegexMatch(A_Loopfield, "([0-9]+$)", FirstNum)
				;FirstNum:= RegExReplace(A_Loopfield, "^[0-9a-zA-Z_]*([0-9]*$)", "$1")
				Num :=1
				}
			continue
			}
		if Num >1
			{
			Outstring := Outstring . FirstVar . FirstNum . "-" . FirstVar . FirstNum+Num-1 . A_Space . A_Loopfield . A_Space
			}
		else
			{
			Outstring := Outstring . A_Loopfield . A_Space
			}
		CurVar :=
		Num :=0
		FirstVar:=
		FirstNum:=
		}

	}
;This is only for the last element
if Num > 0
	Outstring := Outstring . FirstVar . FirstNum . "-" . FirstVar . FirstNum+Num-1
else if ((!RegexMatch(A_Loopfield, "-")>0)&(RegexMatch(A_Loopfield, "^[a-zA-Z].*[0-9]$")>0))
	Outstring := Outstring . FirstVar . FirstNum
else
	Outstring := Outstring . A_Loopfield
Clipboard = %Outstring%
Sendinput {Ctrl down}v{ctrl up}
Return
}



^+r::
{
;
;Press Control Shift R to Repeat sections of syntax. {num:num} and {R} are replaced with a specific number, or use {R+1} (example) to replace with that number plus or minus an integers	
;
Clipboard=
Outstring =
send ^c
Clipwait, 2
if RegexMatch(Clipboard, "`{[\d-+]+:[\d-+]+`}")>0
	{
	InputList := ["`{R`}"]
	ReplaceList := [0]		
	StringReplace, Temp, Clipboard, `r`n, A_Space, All
	StartVal := RegExReplace(Temp, ".*`{([\d-+]+):[\d-+]+`}.*", "$1")
	NumRep := (RegExReplace(Temp, ".*`{[\d-+]+:([\d-+]+)`}.*", "$1")-StartVal+1)
	Temp:=RegExReplace(Clipboard, "{[\d-+]+:[\d-+]+`}", "`{R`}")
	
	if (RegexMatch(Temp, "`{R[\d-+]+`}")>0)
		{
		Loop, parse, Temp, `{R
			{
			if (RegexMatch(A_LoopField, "^[\d-+]+.*")>0)
				{
				x := RegExReplace(A_LoopField, "^([\d-+]+).*", "$1")
				if !hasValue(ReplaceList, x)
					{
					ReplaceList.Insert(x)
					ReplaceText:= "`{R" . x . "`}"
					InputList.Insert(ReplaceText)
					}
				}
			
			}
		
		}

	Loop, %NumRep%
		{
		AddThis:=Temp
		CurIt:=A_Index
		For index, value in InputList
			{
			Repnum:=StartVal+CurIt-1+ReplaceList[index]
			StringReplace, AddThis, AddThis, %value%, %Repnum% , All
			}
		Outstring:=OutString . AddThis . "`r`n"
		}
	Clipboard := Outstring
	Sendinput {Ctrl down}v{ctrl up}
	Return
	}
else if RegexMatch(Clipboard, "`{.+?\s.+?")>0
	{
	candidate_sections := StrSplit(Clipboard, "`}")
	For index, value in candidate_sections
		{
		repeatwords := RegexReplace(value, "^.*?`{", "")
		tmp := RegExReplace(repeatwords, "\b", Replacement = "\b", count)
		if count > 2
			{
			replacelist := StrSplit(repeatwords, A_Space)
			break
			}
		}
	;replacelist := StrSplit(repeatwords, A_Space)

	repeattext :=RegExReplace(Clipboard, "`{.+?`}", "`{R`}")

	For index, value in replacelist
		{
		StringReplace, AddThis, repeattext,`{R`}, %value% , All
		Outstring := OutString . AddThis . "`r`n"
		}
	Clipboard := Outstring
	Sendinput {Ctrl down}v{ctrl up}
	Return
	}
else
	{
	MsgBox No valid syntax found to iterate, should look like: {1:6} to iterate over numerical range, or {replacement1 replacement2 replacement3} to iterate over string items.
	return
	}
}


^+a::
{
send ^c
ClipWait, 2
if (RegexMatch(Clipboard, "[A-Z]")>0)
	StringLower, Clipboard, Clipboard
else
	StringUpper, Clipboard, Clipboard
Sendinput {Ctrl down}v{ctrl up}
Return
}

^!\::
{
;
;Put slashes between results
;
SetTitleMatchMode 2
IfWinNotActive, Excel
	Return
Clipboard=
send {control down}c{control up}
ClipWait, 2
ResultsList =
Text := Clipboard
Loop, parse, Text, `n, `r
	{
	;Msgbox "%A_LoopField%"
	if (A_LoopField="")
		{
		continue
		}
	StringSplit, Sections, A_LoopField, %A_Tab%
	if Sections0 = 2
		{
		if ((RegexMatch(Sections1, "[0-9]")>0)&(RegexMatch(Sections2, "[0-9]")>0))
			ResultsList := ResultsList . Sections1 . "`/" . Sections2 . "`r`n"
		else
			ResultsList := ResultsList . "`r`n"
		}
	if Sections0 = 3
		{
		if ((RegexMatch(Sections1, "[0-9]")>0)&(RegexMatch(Sections2, "[0-9]")>0)&(RegexMatch(Sections3, "[0-9]")>0))
			ResultsList := ResultsList . round(Sections1, 2) . " (" . round(Sections2, 2) . ", " . round(Sections3, 2) . ")`r`n"
		else
			ResultsList := ResultsList . "`r`n"
		}
	}
Clipboard := ResultsList
MsgBox Manually paste results in desired location
Return
}

^!v::
{
;
;Put results into Visio template
;
Clipboard=
send {control down}c{control up}
ClipWait, 2
ResultsList =
Loop, parse, Clipboard, `n`r
	{
	StringSplit, Sections, A_LoopField, %A_Tab%
	;msgbox % Sections0
	if (A_LoopField="")
		continue
	if Sections0 != 2
		continue
	;	{
	;	MsgBox Please select two columns: Coefficients and labels.
	;	Return
	;	}


	if (Sections2 !="")
		ResultsList := ResultsList . Sections1 . A_Tab . Sections2 . "`r`n"
	}
MsgBox % ResultsList
SetTitleMatchMode 2
WinActivate, Visio
WinWaitActive, Visio
Loop, parse, ResultsList, `n, `r
	{
	if (A_LoopField="")
		continue
	StringSplit, Sections, A_LoopField, %A_Tab%
	WinWaitActive, Visio
	Send {control down}f{control up}
	WinWaitActive, Find
	Clipboard=
	Clipboard := Sections2
	ClipWait, 2
	send {del}{control down}v{control up}
	Clipboard=
	sleep, 50
	send {control down}ac{control up}
	if (Clipboard = Sections2)
		{
		Clipboard=
		send {enter}
		}
	else
		{
		send {control down}a{control up}{del}
		Clipboard := Sections2
		ClipWait, 2
		send {control down}v{control up}
		Clipboard=
		send {enter}
		}
	sleep, 100
	send {esc}
	WinWaitActive, Visio
	sleep, 100
	Clipboard := Sections1
	ClipWait, 2
	Sendinput {Ctrl down}v{ctrl up}
	Sleep, 100
	send {esc}
	}
send {esc}
Return
}


^!p::
{
;Batch convert Visio to PDF
StoreClip()
clipboard =
send {alt down}d{alt up}
send {control down}c{control up}
Clipwait, 2
EnvGet, SystemRoot, SystemRoot
Run %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy unrestricted,
WinWait, PowerShell, , 3
clipboard := "cd """ . clipboard . """`r`n$drawings = Get-ChildItem -Filter ""*.vsd""`r`n `r`n $visio = New-Object -ComObject Visio.Application`r`n $visio.Visible = $true`r`n `r`n foreach ($drawing in $drawings)`r`n {`r`n   $pdfname = [IO.Path]::ChangeExtension($drawing.FullName, '.pdf')`r`n   Write-Host ""Converting:"" $drawing.FullName ""to"" $pdfname`r`n   $document = $visio.Documents.Open($drawing.FullName)`r`n$document.ExportAsFixedFormat(1, $pdfname, 1, 0)`r`n }`r`n `r`n finally`r`n {`r`n   if ($visio) `r`n   {`r`n     $visio.Quit()`r`n   }`r`n `r`n }`r`nexit`r`n"
sleep, 100
send {control down}v{control up}
send {return}
return
}

^+"::
{
Clipboard =
send ^c
Clipwait, 2
;Clean up input
StringReplace, Clipboard, Clipboard, `r`n, %A_Space% , All
StringReplace Clipboard, Clipboard, %A_Space%%A_Space%, %A_Space%, All
StringReplace Clipboard, Clipboard, %A_Space%%A_Space%, %A_Space%, All
StringReplace Clipboard, Clipboard, %A_Space%%A_Space%, %A_Space%, All
StringReplace Clipboard, Clipboard, %A_Space%%A_Space%, %A_Space%, All
StringReplace Clipboard, Clipboard, %A_Space%%A_Space%, %A_Space%, All
IfInString, Clipboard, `"
	{
	StringReplace, Clipboard, Clipboard, `",, All
	send {control down}v{control up}
	return
	}
IfInString, Clipboard, `,%A_Space%
	{
	StringReplace, Clipboard, Clipboard, `,%A_Space%, `"`,%A_Space%`", All
	Clipboard = "%Clipboard%"
	send {control down}v{control up}
	return
	}
IfInString, Clipboard, %A_Space%
	{
	StringReplace, Clipboard, Clipboard, %A_Space%, `"`,%A_Space%`", All
	Clipboard = "%Clipboard%"
	send {control down}v{control up}
	return
	}
Return
}

^+l::
{
;
;Label MPlus syntax
;
Clipboard=
send ^c
clipwait, 2
Outline =
if Clipboard Contains `(
	Loop, parse, Clipboard, `n, `r
		{
		If A_Loopfield Contains `(
			{
			label := RegExReplace(A_Loopfield, ".*\((.*)\);", "$1")
			if (RegexMatch(label, "\d")>0)
				{
				label := RegExReplace(label, "\d", "")
				Outline := Outline . RegExReplace(A_Loopfield, "(.*\().*", "$1") . label . "`);`r`n"
				continue
				}
			else
				Outline := Outline . RegExReplace(A_Loopfield, " \(.*\)", "") . "`r`n"
			}
		else
			Outline:=Outline . A_Loopfield . "`r`n"
		}
else
	{
	Loop, parse, Clipboard, `n, `r
		{
		if A_Loopfield Contains `;
			{
			If A_Loopfield Contains ON,WITH,BY
				{
				StringReplace, label, A_LoopField, %A_Space%, , All
				StringReplace, label, label,%A_Space%with%A_Space%, W, All
				StringReplace, label, label,%A_Space%on%A_Space%, O, All
				StringReplace, label, label,%A_Space%by%A_Space%, B, All
				StringReplace, label, label, `;, , All
				StringReplace, label, label, WITH, W, All
				StringReplace, label, label, ON, O, All
				StringReplace, label, label, BY, B, All
				Outline:= Outline . StrReplace(A_LoopField, "`;", "") . " (" . label . ");`r`n"
				}
			else if A_Loopfield Contains [
				Outline:= Outline . StrReplace(A_LoopField, "`;", "") . RegExReplace(A_Loopfield, "\[(.*)\];", " (m$1);`r`n")
			else
				Outline:= Outline . StrReplace(A_LoopField, "`;", "") . " (v" . StrReplace(A_LoopField, "`;", "") . ");`r`n"
			}
		else
			Outline:=Outline . A_Loopfield . "`r`n"
		}
	}
Clipboard := Outline
send {control down}v{control up}
return
}


pStars()
{
Outstring =
Clipboard=
Send ^c
ClipWait, 2
Text := Clipboard
Loop, parse, Text, `n, `r
	{
	StringSplit, Sections, A_LoopField, %A_Tab%
	if (((RegexMatch(A_Loopfield, "[a-zA-Z_]")>0))||((Sections1 = "")||(Sections2 = "")))
		Outstring:= Outstring . A_LoopField . "`r`n"
	else
		{
		Outstring := Outstring . round(Sections1, 2)
		if ((Sections2 < .08) & (Sections2 > .05))
			Outstring := Outstring . "†"
		if Sections2 <= .05
			Outstring := Outstring . "*"
		if Sections2 < .01
			Outstring := Outstring . "*"
		if Sections2 <= .001
			Outstring := Outstring . "*"
		Outstring:= Outstring . A_Tab . Sections2 . "`r`n"
		}
	}
Clipboard := Outstring
send {control down}v{control up}
return
}

deleteColumns(maxColumns, columnHeaders)
{
SetTitleMatchMode 2
IfWinActive, Excel
	{
	send {control down}{home}{control up}
	loop, %maxColumns%
		send {right}
	loop, %maxColumns%
		{
		send {control down}c{control up}
		If Clipboard Contains %columnHeaders%
			{
			send {control down}{space}{control up}
			send {control down}-{control up}
			}
		send {left}
		}
	}
return
}

findColumn(whichColumn)
{
SetTitleMatchMode 2
IfWinActive, Excel
{
	send {control down}{home}{control up}
	Clipboard := whichColumn
	send {control down}f{control up}
	WinWaitActive, Find
	send {control down}v{control up}
	send !f{esc}
	WinWaitActive, Excel
}
return
}

decimalColumns(whichColumn, incDec)
{
SetTitleMatchMode 2
IfWinActive, Excel
{
	IfNotInString, whichColumn, ipse
		{
		send {control down}{home}{control up}
		Clipboard := whichColumn
		send {control down}f{control up}
		WinWaitActive, Find
		send {control down}v{control up}
		send !f{esc}
		}
	WinWaitActive, Excel
	send {down}{control down}{shift down}{down}{shift up}{control up}
	if incDec < 0
		{
		loop % abs(incDec)
			{
			send {alt down}h9{alt up}
			}
		}
	else
		{
		loop , %incDec%
			send {alt down}h0{alt up}
		}
}
return
}


;;Replace spaces and tabs with commas or vice versa, e.g. to copy a list of variable names from functions that require a comma-separated list of parameters to functions that require space-separated variable lists
^+,::
{
	send {control down}c{control up}
	IfInString, Clipboard, `,
		{
		StringReplace, Clipboard, Clipboard, `,, %A_Space%, All
		}
	else
		{
		StringReplace, Clipboard, Clipboard, %A_Space%, `,%A_Space%, All
		StringReplace, Clipboard, Clipboard, %A_Tab%, `,%A_Space%, All
		}
	send {control down}v{control up}
	Return
}

;;Restart the program
^!q::
	{
	Reload
	return
	}

;^+c::
;{
;run, E:\Constrain Hotkey\COnstrain.AHK
;ExitApp
;}


^+m::MsgBox, % (MplusMode := !MplusMode) ? "MPlus mode ON." : "MPlus mode OFF."
#If MplusMode

#if


^!a::
{
;
;Replace slashed range with average
;
Clipboard=
send {control down}c{control up}
ClipWait, 2
ResultsList =
Average :=1
StringReplace, Text, Clipboard, †,, All
IfInString, Text, `(
	{
	MsgBox Use results with significance asterisks, not confidence intervals.
	Return
	}
Loop, parse, Text, `n, `r
	{
	Average :=1
	IfInString, A_LoopField, /
		{
		StringSplit, Sections, A_LoopField, /
		StringReplace, Result1, Sections1, *,, All, UseErrorLevel
		Stars1 := ErrorLevel
		StringReplace, Result2, Sections2, *,, All, UseErrorLevel
		Stars2 := ErrorLevel
		if (Stars1!=Stars2)
			Average :=0
		if (abs(Result1-Result2) >= .1)
			Average :=0
		if Average
			{
			Result :=round((Result1+Result2)/2, 2)
			Stars=
			NumStars:= round((Stars1+Stars2)/2)
			Loop, %NumStars%
				Stars := Stars . "*"
			ResultsList := ResultsList . Result . Stars . "`r`n"
			}
		else
			ResultsList := ResultsList . A_LoopField . "`r`n"
		}
	else
		ResultsList := ResultsList . A_LoopField . "`r`n"
	}
StringReplace, ResultsList, ResultsList, 0., ., All
StringTrimRight, Clipboard, ResultsList, 1
Sendinput {Ctrl down}v{ctrl up}
Return
}






StringJoin(array, delimiter = ";")					;Function to merge strings
{
  Loop
    If Not %array%%A_Index% Or Not t .= (t ? delimiter : "") %array%%A_Index%
      Return t
}

^+\::				;Backslash to forward slash
{
StoreClip()
getSelection()
Outstring := Clipboard
Stringreplace, Outstring, OUtstring, \, x0x0x0, all
Stringreplace, Outstring, OUtstring, /, \, all
Stringreplace, Outstring, OUtstring, x0x0x0, /, all
Clipboard := Outstring
Clipwait, 2
Sendinput {Ctrl down}v{ctrl up}
RestoreClip()
return
}

^+Enter::				;remove line breaks
{
send {control down}c{control up}
Clipboard := removeLinebreaks(Clipboard)
StringReplace, Clipboard, Clipboard, %A_Space%%A_Space%, %A_Space%, All
return
}

removeLinebreaks(lbString)
{
IfInString, lbString, `r`n
	{
	StringReplace, lbString, lbString, `r`n, %A_Space% , All
	}
else
	{
	StringReplace, lbString, lbString, %A_Space% %A_Space%, %A_Space%, All
	StringReplace, lbString, lbString, %A_Space%, `r`n, All
	}
return lbString
}



;^+a::MsgBox, % (AnalMode := !AnalMode) ? "Analysis mode ON." : "Analysis mode OFF."
;#If AnalMode
;	^m::
;	{
;	Results := GetResults("Means")
;	PasteToWin(".out", "Excel", Results, "{down}")
;	return
;	}
;	^i::
;	{
;	Results := GetResults("Intercepts")
;	PasteToWin(".out", "Excel", Results, "{down}")
;	return
;	}
;#if


+^8::
{
pStars()
return
}



GetResults(WhichResult)
{
CopyResults = 0
OutString =
Clipboard = MODEL RESULTS
send {control down}fv{control up}{enter}{shift down}{PgDn 12}{shift up}{control down}c{control up}{esc}
Loop, parse, Clipboard, `n, `r
	{
	If CopyResults
		{
		Line = %A_LoopField%
		StringReplace Line, Line, %A_Space%%A_Space%, %A_Space%, All
		StringReplace Line, Line, %A_Space%%A_Space%, %A_Space%, All
		StringReplace Line, Line, %A_Space%%A_Space%, %A_Space%, All
		StringReplace Line, Line, %A_Space%%A_Space%, %A_Space%, All
		StringReplace Line, Line, %A_Space%%A_Space%, %A_Space%, All
		StringReplace Line, Line, %A_Space%, %A_Tab%, All
;		MsgBox, '%Line%'
		StringSplit, Sections, Line, %A_Tab%
		if Sections0 != 0
			Outstring:= Outstring . Sections1 . A_Tab . Sections2 . A_Tab . Sections5 . "`r`n"
		else
			{
			CopyResults = 0
			Continue
			}
		}
	IfInString, A_LoopField, %WhichResult%
		CopyResults := 1
	IfInString, A_LoopField, STANDARDIZED
		break
	}
;MsgBox, %OutString%

return OutString
}

^+z::
{
;
;Z test script
;
SetTitleMatchMode 2

IfWinNotActive, .inp
	{
	MsgBox Please open an MPlus input file and select a valid list of parameter labels before running this hotkey.
	return
	}
Clipboard=
labelnumber=1
labeldata=
getSelection()
Outstring =
listoflabels=
listofconstraints=

IfInString, ClipBoard, MODEL CONSTRAINT:
	{
	headertext := 1
	labels := Object()
	Loop, parse, Clipboard, `n, `r
		{
		if RegExMatch(A_Loopfield, "L1") = 1
			headertext := 0
		if headertext = 1
			continue
		
		if RegExMatch(A_Loopfield, "L") = 1
			{
			StringReplace, fieldtext, A_LoopField, `;,
			fieldtext := RegExReplace(fieldtext, "^L\d+ = ")
			fieldlabels := StrSplit(fieldtext, "-")
			if !hasValue(labels, fieldlabels[1])
				labels.Insert(fieldlabels[1])
			if !hasValue(labels, fieldlabels[2])
				labels.Insert(fieldlabels[2])
			}
		if (A_Loopfield = "" && labels != "")
			{
			for index, element in labels
				{
				;Msgbox Element %index% is %element%
				Outstring := Outstring . element . A_Space
				}
			StringTrimRight, Outstring, Outstring, 1
			Outstring := Outstring . "`r`n"
			labels := Object()
			}
		}
	if labels != "")
			{
			for index, element in labels
				{
				;Msgbox Element %index% is %element%
				Outstring := Outstring . element . A_Space
				}
			StringTrimRight, Outstring, Outstring, 1
			Outstring := Outstring . "`r`n"
			labels := Object()
			}
	Clipboard =
	sleep, 50
	Clipboard := Outstring
	clipwait, 2
	Sendinput {Ctrl down}v{ctrl up}
	return
	}

Outstring :="MODEL CONSTRAINT:`r`nnew (L1"
Loop, parse, Clipboard, `n, `r
	{
	if A_Loopfield = ""
		Continue

	else if RegExMatch(A_Loopfield, "!") = 1
		{
		listofconstraints := listofconstraints . A_Loopfield . "`r`n"
		Continue
		}

		
	labels := StrSplit(A_Loopfield, [A_Space])
	if labels.MaxIndex() = 1
		{
		labels.Push("0")
		comparelabels := labels[1] . "-" . labels[2]
		listofconstraints := listofconstraints . "L" . labelnumber . " = " . comparelabels . ";`r`n"
		labeldata := labeldata . "L" . labelnumber . A_Tab . comparelabels . "`r`n"
		labelnumber += 1	
		}

	else if labels.MaxIndex() = 2
		{
		comparelabels := labels[1] . "-" . labels[2]
		listofconstraints := listofconstraints . "L" . labelnumber . " = " . comparelabels . ";`r`n"
		labeldata := labeldata . "L" . labelnumber . A_Tab . comparelabels . "`r`n"
		labelnumber += 1	
		}
	else
		{
		for index, element in labels
			{
			curitem := index
			numreps :=  labels.MaxIndex()-Index
			loop, %numreps%
				{
					withitem := curitem+A_Index
					comparelabels := element . "-" . labels[withitem]
					listofconstraints := listofconstraints . "L" . labelnumber . " = " . comparelabels . ";`r`n"
					labeldata := labeldata . "L" . labelnumber . A_Tab . comparelabels . "`r`n"
					;listoflabels := listoflabels . "L" . labelnumber . "`r`n"
					
					labelnumber += 1
				}		
			}
		}
	listofconstraints:=listofconstraints . "`r`n"
	}
	
if labelnumber = 2
	{
	Outstring := Outstring . ");`r`n`r`n!Constraints list:`r`n" . listofconstraints	
	}
else
	{
	Outstring := Outstring . "-L" . labelnumber-1 . ");`r`n`r`n!Constraints list:`r`n" . listofconstraints	
	}

Outstring := Outstring . "`r`n!Begin label and parameter data`r`n"
Loop, Parse, labeldata, `r`n
	{
	if RegExMatch(A_Loopfield, "^L\d") = 1
		Outstring := Outstring . "!" . A_Loopfield . "`r`n"
	}
Outstring := Outstring . "!End label and parameter data`r`n`r`n`r`n"
Clipboard := Outstring
Sendinput {Ctrl down}v{ctrl up}
sleep, 20

;Clipboard := labeldata
;MsgBox Labeldata is now in memory, paste to Excel or read.table("clipboard", sep="\t") in R
RestoreClip()

return
}


^+w::
{
;
;Make wald test script
;
SetTitleMatchMode 2
Clipboard=
Outstring=
labels := Object()
send ^c
clipwait, 2
Loop, parse, Clipboard, `n, `r
	{
	If A_Loopfield Contains `(
			{
			label := RegExReplace(A_Loopfield, ".*\((.*)\);", "$1")
			if !hasValue(labels, label)
				labels.Insert(label)
			}
	}
for index, element in labels
	{
    numreps:=labels.MaxIndex()-index
	Loop, %numreps%
		{
		Outstring:=Outstring . "!0 = " . labels[index] . "-" . labels[index+A_Index] . ";`r`n"
		}
	}
Outstring := Outstring . "!x"
Clipboard=
send ^a^c
Clipwait, 2
if Clipboard Contains MODEL TEST:
	{
	send ^f
	WinWaitActive, Find
	Clipboard = !x
	Sendinput {Ctrl down}v{ctrl up}
	send {return}
	WinWaitActive, .inp
	Outstring := "!r`r`n" . Outstring
	
	Clipboard := Outstring
	Sendinput {Ctrl down}v{ctrl up}
	}
else
	{
	if Clipboard Contains OUTPUT:
		{
		send ^f
		WinWaitActive, Find
		Clipboard = OUTPUT:
		Sendinput {Ctrl down}v{ctrl up}
		send {return}
		WinWaitActive, .inp
		Outstring := "MODEL TEST:`r`n!Start wald`r`n" . Outstring . "`r`nOUTPUT:"
		Clipboard := Outstring
		Sendinput {Ctrl down}v{ctrl up}
		}
	else
		{
		send {PgDn 20}{down}
		Outstring := "MODEL TEST:`r`n!Start wald`r`n" . Outstring . "`r`n"
		Clipboard := Outstring
		Sendinput {Ctrl down}v{ctrl up}
		}	
	}
Clipboard := Outstring
return
}

+^t::
{
;
;Run batch of Wald tests
;
	Waldstring =
	SetTitleMatchMode 2
	IfWinNotActive, .inp
		{
		MsgBox, No valid MPlus input file detected!
		return
		}

	send {control down}f{control up}
	WinWaitActive, Find
	Clipboard = !Start wald
	send {control down}v{control up}{enter}
	WinWaitActive, .inp

	loop
		{
		Send {home}{down}{shift down}{end}{shift up}{control down}c{control up}
		Sleep, 100
		;This one should skip a line if it's empty or nonsense

		if RegexMatch(Clipboard, "^!0")>0
			{
			WaldList := SubStr(Clipboard, 6, -1)
			send {home}{del}
			}
		else if RegexMatch(Clipboard, "^!r")>0
			{
			Waldstring:= Waldstring . "`r`n"
			Continue
			}
		else if RegexMatch(Clipboard, "^!x")>0
			break
		else
			{
			Send {home}
			Continue
			}
		
		Send {alt down}r{alt up}
		Sleep, 10
		SetTitleMatchMode 3
		WinWaitActive, Mplus
			Send {enter}
		SetTitleMatchMode 2
		WinWaitActive, .out
		Stats := GetWald()
		pvalue := RegExReplace(Stats, ".*p = \d?(\.\d{1,2}).*", "$1")
		if(pvalue<.05)
			Waldstring:= Waldstring . WaldList . A_Tab . Stats . A_Tab . "SIG`r`n"
		else
			Waldstring:= Waldstring . WaldList . A_Tab . Stats . "`r`n"
		Send !fc
		WinWaitActive, .inp
		Send {!}
		}
	PasteToWin(".inp", "Excel", Waldstring, "{right}")
return
}

PasteToWin(Referrer, WindowTitle, Text, MoveCur)
{
SetTitleMatchMode, 2
IfWinExist %WindowTitle%
	{
	WinActivate %WindowTitle%
	WinWaitActive, %WindowTitle%
	Clipboard := Text
	send {control down}v{control up}
	Sleep, 50
	Send %MoveCur%
	WinActivate %Referrer%
	WinWaitActive, %Referrer%
	}
else
	MsgBox, Please make sure %WindowTitle% is running!
return
}


+^e::
{
SetTitleMatchMode 2

IfWinNotActive, .out
	{
	IfWinActive, Excel
		{
		Clipboard=
		send ^c
		ClipWait, 2
		Clipboard:=RegExReplace(Clipboard, "\s\([\d\.]+\)")
		;MsgBox % Clipboard
		send {control down}v{control up}
		Return
		}
	else
	{
	MsgBox Please open an MPlus Output file and select a valid output section (coefficients)
	return
	}
	}
Clipboard =
send {control down}c{control up}
Clipwait, 2

Outstring =
ResultsList := Object()
StringReplace, Clipboard, Clipboard, Group, Group, UseErrorLevel
numGroups := ErrorLevel
if(numGroups = 0)
	numGroups := 1

curGroup := 0

StringReplace Clipboard, Clipboard, %A_Space%    %A_Space%, %A_Space%, All
StringReplace Clipboard, Clipboard, %A_Space% %A_Space%, %A_Space%, All
StringReplace Clipboard, Clipboard, %A_Space%%A_Space%, %A_Space%, All
StringReplace Clipboard, Clipboard, %A_Space%%A_Space%, %A_Space%, All
StringReplace Clipboard, Clipboard, %A_Space%%A_Space%, %A_Space%, All
StringReplace Clipboard, Clipboard, %A_Space%, %A_Tab%, All

Loop, parse, Clipboard, `n, `r
	{
	IfInString, A_Loopfield, Group
		{
		if((curGroup = 0)&(numGroups > 1))
			{
			Outstring := A_Loopfield . "`r`n"
			curGroup := curGroup +1
			}
		else
			{
			ResultsList.Insert(Outstring)
			Outstring := A_Loopfield . "`r`n"
			}
		}
	else
		{
		StringSplit, Sections, A_LoopField, %A_Tab%, %A_Space%
		;MsgBox % Sections0
		if (Sections0 = 0)
			{
			Outstring:= Outstring . "`r`n"
			}
		if (Sections0 = 2)
			{
			StringTrimRight, Sections2, Sections2, 1
			Prefix := Sections2 . A_Space
			}
		if ((Sections1 = "Effects")&(Sections2 = "from"))
			{
			Prefix := Sections3 . " to " . Sections5 . A_Space
			}
		if (Sections0 = 3)
			{
			Prefix := Sections2 . A_Space . Sections3 . A_Space
			}
		if (Sections0 = 5)
			{
			Prefix := Sections2 . A_Space . Sections3 . A_Space . Sections4 . A_Space . Sections5 . A_Space
			}
		if (Sections0 = 6)
			{
			Sections3 := round(Sections3, 2) . " `(" . round(Sections4, 2) . "`)"
			Outstring:= Outstring . Prefix . Sections2 . A_Tab . Sections3 . A_Tab . round(Sections6, 3) . "`r`n"
			}
		if (Sections0 = 9)
			{
			Outstring:= Outstring . Prefix . Sections2 . A_Tab . round(Sections6, 2) . A_Tab . round(Sections4, 2) . A_Tab . round(Sections8, 2) . "`r`n"
			}
		if (Sections0 = 7 || (Sections0 = 8 && Sections8 = "*")) 
			{
			Outstring:= Outstring . Prefix . Sections2 . A_Tab . round(Sections3, 2) . A_Tab . round(Sections6, 2) . A_Tab . round(Sections7, 2)
			if Sections0 = 8
				Outstring:= Outstring . A_Tab . "*`r`n"
			else
				Outstring:= Outstring . A_Tab . "`r`n"
			}
		}
	
	}

ResultsList.Push(Outstring)

SetTitleMatchMode, 2
IfWinExist Excel
	{
	WinActivate ; Activate Excel
	}
else
	{
	MsgBox Please open Excel before continuing.
	WinWait Excel
	WinActivate
	}
	
for resultnum, resultsection in ResultsList
	{

	Clipboard =
	Clipboard := resultsection
	Clipwait, 2
	send {control down}v{control up}
	sleep, 50
	send {right 4}
	}
;
;
;



;send {control down}v{control up}
;sleep, 50
;Send !{tab}
return

}



^!s::MsgBox, % (StatsToggle := !StatsToggle) ? "Statistics mode ON." : "Statistics mode OFF."

#If StatsToggle
;Descriptives
$m::Send {control down}i{control up}M{control down}i{control up} = `
$s::Send {control down}i{control up}SD{control down}i{control up} = `
;Test statistics
$F::Send {control down}i{control up}F{control down}i{control up}(
$t::Send {control down}i{control up}t{control down}i{control up}(
$n::Send {control down}i{control up}N{control down}i{control up} = `
$p::Send {control down}i{control up}p{control down}i{control up} = `
^p::Send {control down}i{control up}p{control down}i{control up} < `
!p::Send {control down}i{control up}p{control down}i{control up} > `
$r::Send {control down}i{control up}r{control down}i{control up} = `
+$r::Send {control down}i{control up}R{control down}i{control up}{U+00B2} = `
;Greek letters
$d::Send {U+0394}
$a::Send {U+03B1} = `
$b::Send {U+03B2} = `
$+b::Send B = `
$x::Send {U+03C7}{U+00B2}(
#If

^+q::
SetTitleMatchMode 2
IfWinNotActive, .out
	Send !mr{enter}
WinWaitActive, .out, , 10
Clipboard=
Send {control down}ac{control up}
Clipwait, 2

IfNotInString, Clipboard, TERMINATED NORMALLY
	{
	MsgBox, 4,, WARNING! Please check the output manually for errors. Would you like to continue anyway?
		IfMsgBox No
			{
			Clipboard = WARNING
			send {control down}fv{control up}{enter}
			Exit
			}
	}
WinGetTitle, Title, A
StringReplace, Title, Title, Mplus - [, , All
StringReplace, Title, Title, .out], , All
Outstring := Title . "`r`n"

Clipboard = model fit information
send ^f
WinWaitActive, Find, , 10
Sendinput {Ctrl down}v{ctrl up}
send {enter}
WinWaitActive, .out, , 10
Send {shift down}{PgDn}{PgDn}{PgDn}{shift up}
Clipboard=
send {control down}c{control up}
Clipwait, 2
Results := Clipboard
send {esc}
StringReplace, Results, Results, P-Value, pwaarde, All
StringReplace, Results, Results, Chi-Square Test of Model Fit for the Baseline, onzin, All
StringReplace, Results, Results, Chi-Square Contributions, onzin, All
StringReplace, Results, Results, Chi-Square Test of Model Fit, Chikwadraat, All
StringReplace, Results, Results, H0 Value, onzin, All
;StringReplace, Results, Results, chi-square value, onzin, All
StringReplace, Results, Results, H1 Value, onzin, All
StringReplace, Results, Results, H0 Scaling, onzin, All
StringReplace, Results, Results, H1 Scaling, onzin, All
StringReplace, Results, Results, Sample-Size, sample size, All
Sleep, 50
Loop, parse, Results, `n, `r
{
IfInString, A_LoopField, MODEL RESULTS, break
IfInString, A_LoopField, Number of Free Parameters
	OutString:= Outstring . "Free" . A_Tab . NumExtract(A_LoopField) . "`r`n"
IfInString, A_LoopField, AIC
	OutString:= Outstring . "AIC" . A_Tab . NumExtract(A_LoopField) . "`r`n"
IfInString, A_LoopField, Sample Size Adjusted BIC
	OutString:= Outstring . "Adj BIC" . A_Tab . NumExtract(A_LoopField) . "`r`n"

IfInString, A_LoopField, Chikwadraat
	ChiS := A_Index
IfInString, A_LoopField, SRMR
	Srmr := A_Index
IfInString, A_LoopField, Scaling Correction
	OutString:= Outstring . "scf" . A_Tab . NumExtract(A_LoopField) . "`r`n"
IfInString, A_LoopField, Estimate
	OutString:= Outstring . "RMSEA" . A_Tab . NumExtract(A_LoopField) . "`r`n"
IfInString, A_LoopField, CFI%A_Space%
	OutString:= Outstring . "CFI" . A_Tab . NumExtract(A_LoopField) . "`r`n"
IfInString, A_LoopField, TLI%A_Space%
	OutString:= Outstring . "TLI" . A_Tab . NumExtract(A_LoopField) . "`r`n"

IfInString, A_LoopField, Value
	{
	if (A_Index-Srmr = 2)
		{
		OutString:= Outstring . "SRMR" . A_Tab . NumExtract(A_LoopField) . "`r`n"
		Srmr := 0
		}
	if (A_Index-ChiS = 2)
		OutString:= Outstring . "Chi2" . A_Tab . NumExtract(A_LoopField) . "`r`n"
	}
IfInString, A_LoopField, Degrees of Freedom
	{
	if (A_Index-ChiS = 3)
		OutString:= Outstring . "df" . A_Tab . NumExtract(A_LoopField) . "`r`n"
	}
IfInString, A_LoopField, pwaarde
	{
	if (A_Index-ChiS = 4)
		OutString:= Outstring . "p" . A_Tab . NumExtract(A_LoopField) . "`r`n"
	}

}
Sleep, 20
Clipboard := Outstring
SetTitleMatchMode, 2
IfWinExist Excel
{
WinActivate ; Activate Excel
}
else
{
MsgBox Please open Excel before continuing.
WinWait Excel
WinActivate
}
send {control down}v{control up}{enter} ; paste and hit enter to go to the next line
Send !{tab} ; switch back to original window
return


^!n::
send {control down}c{control up}
Clipboard := NumExtract(Clipboard)
send {control down}v{control up}
Return


NumExtract(String)
{
ExNr := RegExReplace(String, "[a-zA-Z\s\*\(\)]*")
return ExNr
}

NameExtract(String)
{
i := 1
while (i < (StrLen(String)+1))
{
	StringMid, Cur, String, i, 1
	if Cur is alpha
		ExNm := ExNm . Cur
	++i
}
return ExNm
}

WhichWald()
{
send {control down}f{control up}
WinWaitActive, Find
Clipboard = MODEL TEST
send {control down}v{control up}{enter}
WinWaitActive, .out
Send {shift down}{PgDn}{PgDn}{PgDn}{shift up}
Clipboard := ""
while (Clipboard = "")
	send {control down}c{control up}
send {esc}
ReturnWald =
Loop, parse, Clipboard, `n, `r
	{
	IfInString, A_LoopField, 0 =
		{
		IfNotInString, A_LoopField, !
			{
			StringReplace, ReturnWald, A_LoopField, 0 =, , All
			StringReplace, ReturnWald, ReturnWald, `;
			StringReplace, ReturnWald, ReturnWald, `(, , All
			StringReplace, ReturnWald, ReturnWald, `), , All
			StringReplace, ReturnWald, ReturnWald, %A_Space%, , All
			}
		}
	}
return ReturnWald
}

GetWald()
{
	WaldStats =
	send {control down}f{control up}
	WinWaitActive, Find
	Clipboard = Wald Test of Parameter Constraints
	send {control down}v{control up}{enter}
	WinWaitActive, .out
	Send {shift down}{Down}{Down}{Down}{Down}{Down}{shift up}
	Clipboard := ""
	while (Clipboard = "")
		send {control down}c{control up}
	send {esc}
	StringReplace, Clipboard, Clipboard, P-Value, pwaarde, All
	Loop, parse, Clipboard, `n, `r
		{
		IfInString, A_LoopField, Value
			Wald := NumExtract(A_LoopField)
		IfInString, A_LoopField, Degrees
			WaldStats := WaldStats . "χ²(" . NumExtract(A_LoopField) . ") = " . Wald . ", p = "
		IfInString, A_LoopField, pwaarde
			WaldStats := WaldStats . Round(NumExtract(A_LoopField), 3)
			;. "`r`n"
		}

Return WaldStats
}


