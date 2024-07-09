#SingleInstance Force
SetWorkingDir(A_ScriptDir)
SetTitleMatchMode("RegEx")
#Requires AutoHotkey v2.0
Persistent
;===============================================================================
; This is the WayText application code. By Kunkel321. Updated: 7-9-2024	
; New versions on GitHub: https://github.com/kunkel321/WayText
; On AutoHotkey forum: https://www.autohotkey.com/boards/viewtopic.php?f=83&t=129466			
; Optimized for web-entry of third person narratives.
; Uses ini files in an unorthodox way, as quasi databases.
; The script file is not a "stand alone" tool -- the folder and ini files are needed. 
; There are not many user options in the code... They are mostly in the Settings.ini file.
; The WayText application (this), should be used in conjunction with the wtSettings script.
; At the bottom is the InputBuffer Class by Descolada.  See bottom for information.
;===============================================================================

; ^esc::ExitApp ; <--- Ctrl+Esc is emergency kill switch, when debugging.
TraySetIcon("shell32.dll","81")   ; The old Windows "disc defrag" icon.

settingsFile := "wtFiles\Settings.ini" ; Don't change.
If not FileExist(settingsFile) {
	Msgbox(
		"We must exit because the Settings.ini file was not found at location:`n`n" 
		A_ScriptDir settingsFile, "WAYTEXT APPLICATION", 48)
	ExitApp
}

; ===== Additional options, not in wtSettings GUI ========================== 
markerX := "█" ;"►" ; points to fosussed control.  Can't store this in the Settings.ini, because it's a Unicode character.
markerColor := "c0078D7" ; Set as "" for default color. "c0078D7" is the default blue for selected text in Windows 10.
skipNum := 5 ; Navigate the Definitions ListBox by jump-skipping this many at a time. (Shift+Up/Down)
; ==========================================================================

; This function opens the TextDefinitions.ini file and attempts to navigate to whatever [Section]
; was selected in WayText. 
goToSection(*)
{	ErrorLevel := "ERROR"
	Try { 
		ErrorLevel := Run(varFaveEditor " `"" A_ScriptDir "\wtFiles\TextDefinitions.ini`"", , "", )
		WinWaitActive("TextDefinitions")
		;MsgBox "now we look for: " wtListBoxDefinition.Text
		Sleep 100
		SendInput "^f"
		Sleep 100
		SendInput "[" wtListBoxDefinition.Text "]"
		; -----------------------------------------------
		;Next steps are for Steve's custom keyboard shortcuts in EditPad Pro...  The ini Sections are 
		;individually manually folded, so fold them all, then unfold the one I want to see. 
		Sleep 200 
		SendInput "!^{F3}" ; custom changed from ctrl+F3.   <---- Only for Steve's setup in EditPadPro
		Sleep 100
		SendInput "^+{F8}" ; custom assigned to Fold All.   <---- Only for Steve's setup in EditPadPro
		Sleep 100
		SendInput "!^+{F8}" ; custom assigned to Unfold.   <---- Only for Steve's setup in EditPadPro
		; -----------------------------------------------
	}
	If ErrorLevel {
		Run("wtFiles\TextDefinitions.ini") ; Just runs it with default editor.  Windows might prompt for editor. 
		Sleep 100
		SendInput "^f"
		Sleep 100
		SendInput "[" wtListBoxDefinition.Text "]"
	}
}

#HotIf WinActive("WayText Application",) ; Can't use A_Var here.
^s:: ; When you press Ctrl+s, this scriptlet will save the file, then reload it to RAM.
{	Send("^s") ; Save me.
	;MsgBox("Reloading...", "", "T0.3")
	Sleep(500)
	Reload() ; Reload me too.
	;MsgBox("I'm reloaded.") ; Pops up then disappears super-quickly because of the reload.
}
#HotIf

; Startup/reload anouncement.
SoundBeep(1300, 150)
SoundBeep(1500, 100)
SoundBeep(1500, 100)

;==========Read from Settings.ini file========
; The fourth parameter (when present) is the default value.  Change it desired.
varVisualMode := IniRead(settingsFile, "MainSettings", "VisualSendMode", "0")
useBuffer := IniRead(settingsFile, "MainSettings", "UseInputBuffer", "1")
postPasteDelay := IniRead(settingsFile, "MainSettings", "PostPasteDelay", "200")
varFaveEditor := IniRead(settingsFile, "MainSettings", "FaveEditor", "SciTE\SciTE.exe")
varTargetWins := IniRead(settingsFile, "MainSettings", "PreferredTargetWins", "Notepad")
varForcePref := IniRead(settingsFile, "MainSettings", "ForceUseOfPreferred", "0")
varHotKey := IniRead(settingsFile, "MainSettings", "MyHotkey", "!+w")
varTodayAt := IniRead(settingsFile, "Offedule", A_DDDD, "1")
varColorNormal := IniRead(settingsFile, "MainSettings", "GUIcolor", "E5E4E2")
varColorWarning := IniRead(settingsFile, "MainSettings", "GUIwarnColor", "FC8EAC")
varListColor := IniRead(settingsFile, "MainSettings", "ListColor", "Default")
varFontColor := IniRead(settingsFile, "MainSettings", "FontColor", "Default")
varDefaultMenuColor := IniRead(settingsFile, "MainSettings", "MenuColorIsDefault", "1")
varMemoryMinutes := IniRead(settingsFile, "MainSettings", "MemoryMinutes", "60")
varDefName  := IniRead(settingsFile, "MainSettings", "defaultName", "Client")
;=========================================
menuBackColor := SubStr(IniRead(settingsFile, "MainSettings", "GUIcolor", "E5E4E2"), -6) ; This sets the background
; color of the systray icon menu.  We can't change the font color, so if you opt for a dark gui with light font, 
; just set it to menuBackColor := "Default"

LastRunTime := 0 ; Used for "remember last used" code, below. 

; Convert custom hotkey to human-friendly format for display in top of menu. 
hkVerbose :=  (inStr(varHotKey, "^")?"Ctrl+":"") (inStr(varHotKey, "+")?"Shift+":"") 
		. (inStr(varHotKey, "!")?"Alt+":"") (inStr(varHotKey, "#")?"Win+":"") (StrUpper(SubStr(varHotKey, -1)))

; ------- build menu and submenu -------------
wayTxMenu := A_TrayMenu ; Tells script to use this when right-click system tray icon.
wayTxSubMenu := Menu() 
wayTxMenu.Delete ; Removes all of the defalt memu items, so we can add our own. 
menuBackColor := (varDefaultMenuColor = "1")? "Default" : SubStr(varColorNormal, -6)
wayTxMenu.SetColor(menuBackColor) 
wayTxSubMenu.SetColor(menuBackColor) 

wayTxMenu.Add(hkVerbose, wtTips)
wayTxMenu.SetIcon(hkVerbose, "icons/wtSubMenu8.ico")
wayTxMenu.Add("Restart WayText", (*) => Reload())
wayTxMenu.SetIcon("Restart WayText", "icons/reload.ico")

myTabs := ["Activation","Send Mode","Editor","Targets","Colors","Schedule","Pre-Fill","Tips"]
; myTabs array must match the one in wtSettings.ahk file.
for tab in myTabs { ; Makes one menu item per item in myTabs array.
	wayTxSubMenu.Add(myTabs[A_Index], ButtonSettings)
	wayTxSubMenu.SetIcon(myTabs[A_Index], "icons/wtSubMenu" A_Index ".ico") 
} ; The ico files must be named correctly for this to work, "wtSubMenu1.ico" etc.

wayTxMenu.Add("Open Settings", wayTxSubMenu) 
wayTxMenu.SetIcon("Open Settings", "icons/gear.ico")
wayTxMenu.Add("Open Definitions INI", ButtonDefinitions)
wayTxMenu.SetIcon("Open Definitions INI", "icons/scroll.ico")
wayTxMenu.Add("Open Offices INI", ButtonOffices)
wayTxMenu.SetIcon("Open Offices INI", "icons/Office.ico")
wayTxMenu.Add("List Lines Debug", (*) => ListLines())
wayTxMenu.SetIcon("List Lines Debug", "icons/list.ico")
wayTxMenu.Add("Exit Script", (*) => ExitApp())
wayTxMenu.SetIcon("Exit Script", "icons/close.ico")
; ----- end of menu creation ----

if (varVisualMode = '1')
	SendMode("Event") 	; Not recommended unless 'use inputbuffer' = true, then yes, use.
else SendMode("Input")	;(Recommended by AHK) faster but you often can't see progress of typing.

varListColorCode := SubStr(varListColor, -6) ; Extracts just the hex code (we don't need the color name).
varFontColorCode := SubStr(varFontColor, -6)

; ---- create gui object ----

wt := Gui(, 'WayText Application') ; Settings app uses this, so don't change name. 
wtHwnd := wt.Hwnd	
wt.Opt("-MinimizeBox +alwaysOnTop")
wt.SetFont("s11 c" varFontColorCode)

; ---- name box ----
NameX := wt.Add("Text", "x2 " markerColor, markerX) ; The "X" variables are all for the visual marker. 
wtEditClientName := wt.Add("Edit", "x14 y8 w220 h25 Background" varListColorCode " vClientName", varDefName)  ; Box to type in name.
NameX.Visible := False
wtEditClientName.OnEvent 'Focus'    , (wtEditClientName, *) => NameX.Visible := True ; These "focus" items are for the markX symbol.  
wtEditClientName.OnEvent 'LoseFocus', (wtEditClientName, *) => NameX.Visible := False

; ---- gender radio group ----
RadioX := wt.Add("Text", "x2 " markerColor, markerX)
wt.SetFont("s10")
radMale := wt.Add("Radio", "x+2 y40 w60 h20 vGender checked1", "&Male")  ; Radio group for gender.
radFemale := wt.Add("Radio", "xp+60 y40 w80 h20 ", "&Female")
radNeutral := wt.Add("Radio", "xp+78 y40 w80 h20 ", "&Neutral")
RadioX.Visible := False
radMale.OnEvent 'LoseFocus'			, (wtEditClientName, *) => RadioX.Visible := False
radFemale.OnEvent 'LoseFocus'		, (wtEditClientName, *) => RadioX.Visible := False
radNeutral.OnEvent 'LoseFocus'		, (wtEditClientName, *) => RadioX.Visible := False
radMale.OnEvent 'Focus'    			, (wtEditClientName, *) => RadioX.Visible := True
radFemale.OnEvent 'Focus'    		, (wtEditClientName, *) => RadioX.Visible := True
radNeutral.OnEvent 'Focus'    		, (wtEditClientName, *) => RadioX.Visible := True
wt.SetFont("s11 ")

; ---- target window label ----
wtLabelTarget := wt.Add("Text", "x240 y6 w138 h108 +wrap", "--- Target Window ---`n " . "Variable with Win target name goes here.")

; ---- Offices list ----
OfficeX := wt.Add("Text", "x2 y72 " markerColor, markerX)
AllOfficesArr := []
AllOffices := ""
Try AllOffices := IniRead("wtFiles\OfficeList.ini") ; Gets `n-delimited list from ini file. 
If (AllOffices = "") ; If file not found, use below error statement.
	AllOffices := "'OfficeList.ini' file not found.`nIt should be in the 'wtFiles' folder."
Loop parse, AllOffices, "`n"
	AllOfficesArr.Push(A_LoopField)
OffRows := AllOfficesArr.Length
wtListBoxOff := wt.Add("ListBox", "r" . OffRows . " xm+2 y73 w220 Background" varListColorCode " vOffChoice Choose" . varTodayAt . " section", AllOfficesArr)
OfficeX.Visible := False
wtListBoxOff.OnEvent 'Focus'    	, (wtListBoxOff, *) => OfficeX.Visible := True
wtListBoxOff.OnEvent 'LoseFocus'	, (wtListBoxOff, *) => OfficeX.Visible := False

; ---- Definitions list ----
DefinitionX := wt.Add("Text", "x2 " markerColor, markerX)	
AllDefinitions := ""
AllDefinitionsArr := []
AllDefinitions := ""
Try AllDefinitions := IniRead("wtFiles\TextDefinitions.ini") ; Gets `n-delimited list from ini file. 
If (AllDefinitions = "") ; If file not found, use below error statement.
	AllDefinitions := "'TextDefinitions.ini' file not found.`nIt should be in the 'wtFiles' subfolder."
Loop Parse, AllDefinitions, "`n"
	AllDefinitionsArr.Push(A_LoopField)
NoteRows := (AllDefinitionsArr.Length > 32)? 32 : AllDefinitionsArr.Length
wtListBoxDefinition := wt.Add("ListBox", "r" . NoteRows . " x+4 w352 Background" varListColorCode " vNoteChoice  Choose1 section", AllDefinitionsArr)
wtListBoxDefinition.OnEvent("DoubleClick", ButtonInsert) ; Doubleclicking list item enters it (no need for Insert Button).
DefinitionX.Visible := False
wtListBoxDefinition.OnEvent 'Focus'    	, (wtListBoxDefinition, *) => DefinitionX.Visible := True
wtListBoxDefinition.OnEvent 'LoseFocus'	, (wtListBoxDefinition, *) => DefinitionX.Visible := False

; ---- buttons ----
wtButtonInsert := wt.Add("Button", "w80 h30 Section", "Insert").OnEvent("Click", ButtonInsert)
wtButtonDefinitions := wt.Add("Button", "w95 ys x+5  h30", "Definitions")
	wtButtonDefinitions.OnEvent("Click", ButtonDefinitions) ; Normal left-click opens definitions ini. 
	wtButtonDefinitions.OnEvent("ContextMenu", goToSection) ; Right-click opens ini then attempts to go to [Section]. 
wtButtonSettings := wt.Add("Button", "w80 ys x+5  h30", "Settings")
	wtButtonSettings.OnEvent("Click", ButtonSettings) ; Normal left-click opens settings. 
	wtButtonSettings.OnEvent("ContextMenu", (*) => wayTxSubMenu.Show()) ; Right-click opens menu to go to specific tab in settings. 
wtButtonCancel := wt.Add("Button", "w80 ys x+5  h30", "Cancel").OnEvent("Click", ButtonCancel)
; ---- end of GUI creation ----

; ---- context-specific hotkeys allow selection to "loop" through or "skip/jump" in Definitions listbox ----
#HotIf WinActive(wtHwnd) and wtListBoxDefinition.Focused and wtListBoxDefinition.Value = AllDefinitionsArr.Length
	Down::wtListBoxDefinition.Value := 1
#HotIf WinActive(wtHwnd) and wtListBoxDefinition.Focused and (wtListBoxDefinition.Value <= AllDefinitionsArr.Length - skipNum) 
	+Down::wtListBoxDefinition.Value := wtListBoxDefinition.Value + skipNum
#HotIf WinActive(wtHwnd) and wtListBoxDefinition.Focused and wtListBoxDefinition.Value = 1
	Up::wtListBoxDefinition.Value := AllDefinitionsArr.Length
#HotIf WinActive(wtHwnd) and wtListBoxDefinition.Focused and (wtListBoxDefinition.Value >= skipNum) 
	+Up::wtListBoxDefinition.Value := wtListBoxDefinition.Value - skipNum

#HotIf WinActive(wtHwnd) ; Only active when wt form is active.
	Enter::ButtonInsert
	!Enter::ButtonInsert ; Holding Alt also activates debug message.
    Esc::ButtonCancel  
	+Left::wtEditClientName.Focus() ; Shift+Left is "Go to Client name box."
	+Right::wtListBoxDefinition.Focus() ; Shift+Right is "Go to Definitions listbox."
	F1::wtTips ; Show tips message box.
#HotIf  ; End of context-sensitive part.

; The main wt GUI form is created when the script starts, but several pieces of information must
; be collected before it is shown.  The active window is determined.  This gets used for the
; Target Window label and also to determine the color of the gui.  The position to show the gui 
; is determined, and the previous name, etc, are applied if appropriate, then the form is shown.
ThisWinTitle := "", targetWindow := ""
Hotkey varHotKey, startHere ; <--- This is the custom hotkey that gets defined in the settings app.
startHere(*)
{	Global targetWindow := WinActive("A")  ; Get the handle of the currently active window
	ThisWinTitle := WinGetTitle("ahk_id " targetWindow)
	wtLabelTarget.text := "--- Target Window ---`n" ThisWinTitle
	If InStr(ThisWinTitle, "Message")  ; An Outlook email
		Global PostPasteDelay := 900
	else Global PostPasteDelay := 200

	If  RegExMatch(ThisWinTitle, '(' varTargetWins ')') { ; Checks if this is a preferred target winodw.
		varColor := varColorNormal ; This is the normal color of the form.
		nonPreferredWin := 0
	}
	else {
		varColor := varColorWarning ; This is the Warning Color, indicating non-preferred Window.
		global nonPreferredWin := 1
	}
	Global varColorCode := SubStr(varColor,- 6) ; Settings color list formatted as "Name - Hexcode."
	wt.BackColor := varColorCode ; Update color of main GUI.
	
	WinGetPos(&X, &Y, &W, &H, "A")   ; "A" to get the active window's pos.
	If ThisWinTitle = "wtSettings" {
		Global Xpos := X - 508 ; Position wt next to settings gui. 
		Global Ypos := Y
	}
	Else {
		Global Xpos := X + (W * 0.05) ; Use these with GUI Show.
		Global Ypos := Y + (H * 0.2)
	}

	If (varMemoryMinutes * 60000) < (A_TickCount - LastRunTime) ; If X seconds since last run, reset defaults.
	{	wtEditClientName.Text := varDefName
		radMale.Value := 1
		wtListBoxOff.Value := varTodayAt
		wtListBoxDefinition.Value := 1
	}

	If wtEditClientName.Text = varDefName
		wtEditClientName.Focus() ; Must set focus after wingetpos. 
	else wtListBoxDefinition.Focus()
	wt.Show("x" . Xpos . " y" . Ypos) ; Position form based on coordinates obtained above.
}

; This function assesses which TextDefinition and Office entries were selected, then reads the
; ini files and gets the associated information.  The first text "Find-and-replacs" are made to the 
; entry.  The keys of the ini section are delimited, after determining a unique delimiter.  
; keys names are checked for "MiniForm" flag, and CallMiniForm() function is called as needed. 
ButtonInsert(*)
{	Global SelGender := ""
	SelGender := wt.Submit() ; Returns 1,2,3, for use below.
	Global MyEntry := ""
	MyEntry := IniRead("wtFiles\TextDefinitions.ini", wtListBoxDefinition.Text,, "Section [" wtListBoxDefinition.Text "] not found...")  ; From Definition [Section], goes back to INI, gets Definition content (section).
	MyContactInfo := IniRead("wtFiles\OfficeList.ini", wtListBoxOff.Text, "ContactInfo", "Contact Info not found.") ; From Office choice, gets contact info.
	MyOffPhone := IniRead("wtFiles\OfficeList.ini", wtListBoxOff.Text, "OffPhone", "Office phone not found.")
	Global varHeShe := "", varHimHer := "", varHisHer := "", varHisHers := ""
	;MsgBox 'just after iniread, myEnt is:`n' MyEntry 
	if (SelGender.gender = 1) ; Checks the results of the gender radio buttons group and assigns variables.
	{	varHeShe := "he"
		varHimHer := "him"
		varHisHer := "his"
		varHisHers := "his"
	}
	if (SelGender.gender = 2)
	{	varHeShe := "she"
		varHimHer := "her"
		varHisHer := "her"
		varHisHers := "hers"
	}
	if (SelGender.gender = 3)
	{	varHeShe := "they"
		varHimHer := "them"
		varHisHer := "their"
		varHisHers := "theirs"
	}

	MyEntry := StrReplace(MyEntry, "[n]", wtEditClientName.text) ; Makes the replacements.
	MyEntry := StrReplace(MyEntry, "[e]", varHeShe) ; Gender pronouns.
	MyEntry := StrReplace(MyEntry, "[m]", varHimHer)
	MyEntry := StrReplace(MyEntry, "[s]", varHisHer)
	MyEntry := StrReplace(MyEntry, "[r]", varHisHers)
	MyEntry := StrReplace(MyEntry, "[c]", MyContactInfo)
	MyEntry := StrReplace(MyEntry, "[p]", MyOffPhone)

	If RegExMatch(MyEntry, "\.\w") ; Dot followed by character suggests URL or file ext present. 
		Global dotIsOK := 1
	else Global dotIsOK := 0

	; ==== Split-Up ini keys from selected ini section =========
	Global EntryArr := [] ; Declare before use. 
	If !InStr(MyEntry, '=') { ; No key=value, just boilerplate text.
		EntryArr.Push({key:"",value:MyEntry})
	}
	Else {
		;del := 'µ'
		for dx in ['|','$','@','¢','¤','¥','¦','§','©','ª','«','®','¶','µ'] ; Find a delimiter that doesn't exist in entry (ini section).
				If !InStr(MyEntry, dx) { ; 14 possibilities :)
					del := dx
					Break
				}
		If InStr(MyEntry, del) {
			MsgBox 'Cannot parse item. All of the possible parsing`ndelimiters are present in the boilerplate text for`nthe item: [' wtListBoxDefinition.Text '].'
			Exit
		}
		loop parse, MyEntry, "`n" {
			If InStr(A_LoopField, "=")
				WithDelims .= del A_LoopField '`n'
			Else WithDelims .= A_LoopField '`n'
		}
		WithDelims := SubStr(WithDelims, StrLen(del)+1) ; Trim off starting delimiter
		Pairs := StrSplit(WithDelims, del)
		for pair in Pairs {
			sPair := StrSplit(pair, '=')
			EntryArr.Push({key:sPair[1],value:Trim(sPair[2], '`n')})
		}
	}

	Global mfListArr := []
	for idx, item in EntryArr {
		If InStr(item.key, "MiniForm") {
			Loop parse, item.value, '`n'
				mfListArr.Push(A_LoopField)
			Local myOptions := CallMiniForm(item.key, mfListArr) 
			;item.key := StrReplace(item.key, "MiniForm", "")
			item.value := myOptions
			mfListArr := []
		}
	}

	Global LastRunTime := A_TickCount
	; for thisVal in EntryArr
	; 	tempComp .= thisVal.value ' '
	; MsgBox 'before other reps`n' tempComp
	OtherReplacements(EntryArr)
}	

; This function only gets used if one of the key names has the word "MiniForm" in it.  
; It gets called once for each.  The key value is parsed into a gui at runtime and
; presented to the user.  Items might be prechecked, or a radio group might be used.
; The resuslt is saved to text, and punctuation added, if needed.  The gui is 
; recreated and destroyed with each function call. 
CallMiniForm(theKey:="", theList:="") 
{	mfArr := []
	Local MyOptions := "", checked := 0
	mfType := InStr(theKey, "radio")? "Radio" : "Checkbox" ; Default is checkbox.  
	mfPuntuate := InStr(theKey, "punct")? "1" : "0" ; Will match "punctuate" or "punctuation."
	mfLabel := RegExReplace(theKey, "i).*\{(.+)\}.*", "$1") ; Extract {label text}.
	mf := Gui(, 'MiniForm') ; "mf" for "MiniForm"
	mf.BackColor := varColorCode
	mf.OnEvent("Close", mfClose)
	mf.Opt("-MinimizeBox +LastFound +alwaysOnTop")
	mf.SetFont("s11 c" varFontColorCode)
	mf.Add("Text", , mfLabel)
	for idx, Item in theList {
		if RegExMatch(Item, "(1|0) .*") { ; Assess if item should be pre-checked. 
			checked := SubStr(Item, 1,1)
			item := SubStr(Item, 3)
		}
		else checked := 0
		itemWidth := ""
		If StrLen(Item) > 100 ; It the string is long, set width of check/radio item. 
			itemWidth := " w600 "
		mfArr.Push( mf.Add(mfType, itemWidth "checked" checked " v" Idx, item) )
	}
	mf.Add("button", "w80 +Default", "OK").OnEvent("click", mfButtOK)
	mf.Show("x" . Xpos+20 . " y" . Ypos+5) 
	mf.OnEvent("Escape", CloseMF)
	CloseMF(*) 
	{	mf.Destroy
		Reload
	}
	WinWaitClose
	Return  MyOptions
	
	mfButtOK(*)
	{	Selected := 0, subSelected := 0, period := ""
		if mfPuntuate = 1 {
			for Item in mfArr {
				;MsgBox 'loop ' A_Index '`nnumber of items ' mfArr.Length '`n' Item.Text ' is ' Item.Value '`n`nselected: ' Selected
				If Item.Value = 1 {
					if SubStr(Item.text, -1) = '.' ; Do the items have periods? 
						period := "."
					Selected++ ; Count how many selected items.
				}
			}
			If Selected > 2	{ ; 3 or more, so use Oxford comma.
				for Item in mfArr {
					If Item.Value = 1 ; If checkbox is checked.
					{	subSelected++
						If subSelected = 1
							myOptions .= Item.Text ; First checked item.
						Else If subSelected > 1 and subSelected < Selected						
							myOptions .= ", " Item.Text ; Middle checked item(s).
						Else
							myOptions .= ", and " Item.Text ; Last checked item. 
					}
					Else ; Not a checked item, so...
						Continue
				} 
				myOptions := StrReplace(myOptions, ".", "") . period ; Remove all the periods, but put one back at the end when appropriate. 
			}
			Else If Selected = 2	{ ; just use 'and'
				for Item in mfArr {
					If Item.Value = 1 {
						subSelected++
						If Selected > subSelected ; don't use a_index.  compare sel with 'thisSel++'
							myOptions := Item.Text 
						Else
							myOptions .= " and " Item.Text
						;MsgBox 'if sel=2`nSel ' Selected '`nloop ' A_Index '`n' MyOptions
					} 
				}
			myOptions := StrReplace(myOptions, ".", "") . period
			}
			Else ; Only one item selected. 
				for Item in mfArr
					If Item.Value = 1 
						myOptions := Item.Text 
		} 
		Else { ; No punctuation needed.
			for Item in mfArr
				If Item.Value = 1
					myOptions .= " " Item.Text
		}
		mf.Destroy
	} ; bottom of mfButtOK nested function
	mfClose(*)
	{	mf.Destroy
	}
} ; bottom of miniform function

; A simple left-click on the Definitions button calls this.  An attempt is made to use
; the "favorite editor", otherwise a generic launch is done. 
ButtonDefinitions(*)
{	ErrorLevel := "ERROR"
	Try ErrorLevel := Run(varFaveEditor " `"" A_ScriptDir "\wtFiles\TextDefinitions.ini`"", , "", )
	If ErrorLevel
		Run("wtFiles\TextDefinitions.ini")
}

; There's not actually a button for this.  It gets called from the systray menu item. 
ButtonOffices(*)
{	ErrorLevel := "ERROR"
	Try ErrorLevel := Run(varFaveEditor " `"" A_ScriptDir "\wtFiles\OfficeList.ini`"", , "", )
	If ErrorLevel
		Run("wtFiles\OfficeList.ini")
}

; The Gui button and the submenu items all call this. It uses the myTabs array that 
; is defined in the section where the menu is created. 
ButtonSettings(MenuItem,*)
{	ItemPos := 0
	for tab in myTabs {
		if tab = MenuItem { ; myTabs is an array of the Tab names in the settings script. 
			ItemPos := A_Index
			Break
		}
	}
	Run "wtSettings.exe /script wtSettings.ahk " ItemPos
}

; Just hides the form.
ButtonCancel(*)
{ 	wt.Hide()
}

; The top item in the systray menu shows the currently assigned hotkey (in human-friendly format)
; Clicking it shows these tips. 
wtTips(*)
{	wtTips := 
	(
		' Main Hotkey: ' hkVerbose
		'`n`n *Other hotkeys (for when form is active) *`n`n'
		'Shift+Up/Down in Definition List does `"skip-jump.`"`n'
		'Skip-jump number set with variable `"skipNum`", above`n`n'
		'Right-Clicking Settings Button Attempts to open TextDefinitions.ini and use Find feature of editor to search for whatever Definition is selected in List.`n`n'
		'Shift+Clicking Insert Button shows debug message of the selected ini Section.`n`n'
		'Shift+Left to make Name box active.`n`n'
		'Shift+Right to make Definitions listbox active.'
	)
	msgbox wtTips, 'WAYTEXT TIPS', 262144+64
}

; This function makes additional "find-and-replaces" to the entry text.  Does clipboard 
; insertion and time stamp as needed.  If (only if) gender is neutral, several grammar
; fixes are made to the entry verbiage to convert binary-to-non-binary (I.e. singular-to-plural).
OtherReplacements(EntryArr) 
{	combinePastes := 0 ; Reset each time a Definition entry is sent.
	for thisVal in EntryArr { ; thisVal = 'this loop's value'
		tVal := thisVal.value
		if InStr(thisVal.key, "paste") ; So that combinPaste function is only called if there are paste keys.
			Global combinePastes := 1

		if InStr(tVal, "{Clipboard}") ; Look for clipboard flag and insert clipboard contents if present.
		{	myClipContent := A_Clipboard
			tVal := StrReplace(tVal, "{Clipboard}", myClipContent)
		}

		if InStr(tVal, "{now")  ; Look for datestamp tag and insert if needed.
		{	RegExMatch(tVal, "{now\,? ?(.+?)}", &DateFormat)
			MyDate := FormatTime(A_Now, DateFormat[1])
			tVal := StrReplace(tVal, DateFormat[0], MyDate)
		}

		if (SelGender.gender = 3) ; This part gets skipped if gender = male or female.
		{	If InStr(tVal, "they ")
			{	tVal := StrReplace(tVal, "they is", "they are") ; Fix gramar if gender neutral.
				tVal := StrReplace(tVal, "they's", "they are") ; Can break with "he's done that."
				tVal := StrReplace(tVal, "they has", "they have")
				tVal := StrReplace(tVal, "they was", "they were")
				segment := ""
				If (RegExMatch(tVal, "\.{2,}") > 0) { ; Multiple sentences, so search each. 
					Loop Parse, tVal, "." { 
						If InStr(A_LoopField, "they") { ; Only search sentences with 'they'.
							MoreRegExes(A_LoopField) ; <----------  First key is getting duplicated :(
							combined .= segment ". "
							tVal := StrReplace(combined, ". .", ".") ; Gets rid of that extra dot that's caused by the last parse loop.
						}
					}
				}
				Else { ; Array element only has one sentence.
					MoreRegExes(tVal) 
					tVal := segment
				}
				MoreRegExes(tempVal) {
				;-----EXPERIMENTAL! will the RegExes below cause grammar mistakes? They are intended to look for
				;plural words that follow "they" and make them singular.  For example "They needs help"
				;is changed to "They need help."  Works well, but is not perfect.  
					tempVal := RegExReplace(tempVal, "s)((T|t)hey\s\w*)ies\b", "$1y") ; 'he parties --> they party'
					tempVal := RegExReplace(tempVal, "s)((T|t)hey\s\w*ly\s\w*(s|z|sh|ch|x|o))es\b", "$1")  ; 'he really parties --> they really party'
					tempVal := RegExReplace(tempVal, "s)((T|t)hey\s\w*(s))es\b", "$1$3") ; Prevents dropping 's' e.g. 'they surpass.'
					tempVal := RegExReplace(tempVal, "s)((T|t)hey\s\w*(z|sh|ch|x|o))es\b", "$1") ; 'he watches --> they watch'
					tempVal := RegExReplace(tempVal, "s)((T|t)hey\s\w*ly\s\w*)s\b", "$1") ; 'he really runs --> they really run'
					segment := RegExReplace(tempVal, "s)((T|t)hey\s\w*)s\b", "$1") ; 'he runs --> they run'
					Return segment
				}
			}
		} ; End of gender grammar fixes.
		thisVal.value := tVal ; Put modified string back into array.
	} ; bottom of for loop.
	ToSending()
	wt.Hide
}

; Windows takes time to execute back-to-back pastes, which can cause erronous results.
; This is partially attenuated by looking for back-to-back "paste keys" and combining 
; them, at runtime, into one big paste key. 
CombineIniKeys(EntryArr) { ; Logic for function from ChatGPT4 :)
	arrayLength := EntryArr.Length
    index := 1
    while (index <= arrayLength) {
        currentItem := EntryArr[index]
        if (InStr(currentItem.key, "paste")) {
			combinedValue := currentItem.value
            nextIndex := index + 1
            while (nextIndex <= arrayLength && InStr(EntryArr[nextIndex].key, "paste")) {
                combinedValue .= " " . EntryArr[nextIndex].value
                EntryArr.RemoveAt(nextIndex) ; Remove the combined paste item
                arrayLength-- ; Adjust the array length after removal
            }
            EntryArr[index].value := combinedValue ; Update the value of the original 'paste' item
        }
        index++
    }
}

; This function is a "pre-launch" for the Sending() function.  It calls the above CombineIniKeys() 
; when needed and also determines if an item should be sent normally, or as a "debug."  
; It also waits for the target window to be active again.  If the target window doesn't become
; active again within four minutes, a "timeout" message is sent. 
ToSending(*)
{	If combinePastes = 1 
		CombineIniKeys(EntryArr)
	Global ClipBrdOld := A_Clipboard ; Store clipboard contents.
	A_Clipboard := '' ; Clear clipboard.
	If GetKeyState('Shift') || ((varForcePref = 1) and (nonPreferredWin = 1)) ; This will be the modifier key to envoke DEBUG MODE.
	{	Debug()
		Return ; Stops script before actually typing out the content. 
		;Reload
	}
	;SoundBeep 1700, 400 ; Announces waiting for target window to be active.
	If WinWaitActive("ahk_id " . targetWindow,,4) {
		;SoundBeep 1400, 400 ; Announces target window is active again.
		Sending()
	}
	Else
		MsgBox 'WinWaitActive timed-out.'
}

; This function (or the Debug() one below), is the last bit of code to get called. 
; Because ini key values often don't retain leading spaces, things can get messed up
; when there are many keys used.  This function tries to "clean up" this issue. 
; Descolada's InputBuffer is (optionally) called at the beginning, then stopped at the 
; end the buffer is stopped and any collected text is sent.  {Sleep} flags are 
; identified and honored.  Each key name is checked for Type vs. Paste and honored.
; Post-paste delayes are enforced, and several variables are cleared for the next use. 
Sending(*)
{	If (useBuffer = 1)
		{	static HSInputBuffer := InputBuffer()
			HSInputBuffer.Start()
		}
	for item in EntryArr 
	{	item.value := StrReplace(item.value, "  ", " ") ; double space --> single space
		item.value := RegExReplace(item.value, "^\w|(?:\.|:)(\s|\R|\Q{Enter}\E)+\K\w", "$U0") ; Makes sure most sentences are capitalized.
		item.value := StrReplace(item.value, "{Enter}", "`n") ; Replaces "{Enter}" with actual whitespace. 
		customSleepTime := 0
		If RegExMatch(item.key, "i).*Sleep\,? ?(\d{1,4}).*", &time) ; Looks for e.g. "{Sleep 350}" in key name and inserts after type/paste. 
			customSleepTime := time[1]
		If (dotIsOK = 0)
			item.value := RegExReplace(item.value, "([a-zA-Z])\.([a-zA-Z])", "$1. $2") ; Puts space after period.. 
		If InStr(item.key, "paste") { ; 'paste' must occur in each key name.
			A_Clipboard := ''
			A_Clipboard := item.value " "
			; MsgBox 'item.val is:`n' item.value
			Send "^v" 
			Sleep(PostPasteDelay)
			;SoundBeep(1200, 500)
		}
		Else 
		{	If not (subStr(item.value, -1)="}")
			{	Send item.value " "
			}
			Else
			{	Send item.value ; Don't want extra space after sending {Tab}
			}
		}
		Sleep customSleepTime ; defined in key name of ini items
	}
	If (useBuffer = 1)
		HSInputBuffer.Stop()
	Sleep '500' ; Give clipboard ample time to change. 
	A_Clipboard := ClipBrdOld ; Put back clipboard contents. 
	Global combinePastes := 0 ; reset after use.	
	Global EntryArr := [] ; Clear after each use.
}

; Under certain situations this function gets called instead of Sending().
; This allows you to see the ini key names without opening the TextDefinitions.ini file. 
; First a MsgBox is shown with the ini key names and values, then an approximation
; of the final output is displayed in a second MsgBox. 
Debug(*)
{	for it in EntryArr {
		If InStr(it.key, "paste") and InStr(it.Value, "{")
			pasteWarn := "**WARNING: Item has {curlies}, but is set to paste.**`n"
		else pasteWarn := ""
		firstDebugMess .= '-- element[' A_Index '] ----------------------------------`nkey:`t|' it.key '|`nval:`t|' it.value '|`n' pasteWarn
	}
	MsgBox '-------- Contents at ToSending() function --------`n' firstDebugMess
	for item in EntryArr {	
		item.value := StrReplace(item.value, "  ", " ") ; double space --> single space
		item.value := RegExReplace(item.value, "^\w|(?:\.|:)(\s|\R|\Q{Enter}\E)+\K\w", "$U0") ; Makes sure most sentences are capitalized.
		item.value := StrReplace(item.value, "{Enter}", "`n") ; Replaces "{Enter}" with actual whitespace. 
		lastDebugMess .= item.value
	}
	MsgBox '-------- Approximation of Final Output --------`n' lastDebugMess
	Global EntryArr := [] ; Clear after each use. 
}

; This stuff below here is the InputBuffer Class by Descolada.  Learn about it here:
;  https://www.autohotkey.com/boards/viewtopic.php?f=83&t=122865
; The comments and code below ---this---line--- were all written by him. 
/**
 * 
 * InputBuffer can be used to buffer user input for keyboard, mouse, or both at once. 
 * The default InputBuffer (via the main class name) is keyboard only, but new instances
 * can be created via InputBuffer().
 * 
 * InputBuffer(keybd := true, mouse := false, timeout := 0)
 *      Creates a new InputBuffer instance. If keybd/mouse arguments are numeric then the default 
 *      InputHook settings are used, and if they are a string then they are used as the Option 
 *      arguments for InputHook and HotKey functions. Timeout can optionally be provided to call
 *      InputBuffer.Stop() automatically after the specified amount of milliseconds (as a failsafe).
 * 
 * InputBuffer.Start()               => initiates capturing input
 * InputBuffer.Release()             => releases buffered input and continues capturing input
 * InputBuffer.Stop(release := true) => releases buffered input and then stops capturing input
 * InputBuffer.ActiveCount           => current number of Start() calls
 *                                      Capturing will stop only when this falls to 0 (Stop() decrements it by 1)
 * InputBuffer.SendLevel             => SendLevel of the InputHook
 *                                      InputBuffers default capturing SendLevel is A_SendLevel+2, 
 *                                      and key release SendLevel is A_SendLevel+1.
 * InputBuffer.IsReleasing           => whether Release() is currently in action
 * InputBuffer.Buffer                => current buffered input in an array
 * 
 * Notes:
 * * Mouse input can't be buffered while AHK is doing something uninterruptible (eg busy with Send)
 */
class InputBuffer {
    Buffer := [], SendLevel := A_SendLevel + 2, ActiveCount := 0, IsReleasing := 0, MouseButtons := ["LButton", "RButton", "MButton", "XButton1", "XButton2", "WheelUp", "WheelDown"]
    static __New() => this.DefineProp("Default", {value:InputBuffer()})
    static __Get(Name, Params) => this.Default.%Name%
    static __Set(Name, Params, Value) => this.Default.%Name% := Value
    static __Call(Name, Params) => this.Default.%Name%(Params*)
    __New(keybd := true, mouse := false, timeout := 0) {
        if !keybd && !mouse
            throw Error("At least one input type must be specified")
        this.Timeout := timeout
        this.Keybd := keybd, this.Mouse := mouse
        if keybd {
            if keybd is String {
                if RegExMatch(keybd, "i)I *(\d+)", &lvl)
                    this.SendLevel := Integer(lvl[1])
            }
            this.InputHook := InputHook(keybd is String ? keybd : "I" (this.SendLevel) " L0 *")
            this.InputHook.NotifyNonText  := true
            this.InputHook.VisibleNonText := false
            this.InputHook.OnKeyDown      := this.BufferKey.Bind(this,,,, "Down")
            this.InputHook.OnKeyUp        := this.BufferKey.Bind(this,,,, "Up")
            this.InputHook.KeyOpt("{All}", "N S")
        }
        this.HotIfIsActive := this.GetActiveCount.Bind(this)
    }
    BufferMouse(ThisHotkey, Opts := "") {
        savedCoordMode := A_CoordModeMouse, CoordMode("Mouse", "Screen")
        MouseGetPos(&X, &Y)
        ThisHotkey := StrReplace(ThisHotkey, "Button")
        this.Buffer.Push(Format("{Click {1} {2} {3} {4}}", X, Y, ThisHotkey, Opts))
        CoordMode("Mouse", savedCoordMode)
    }
    BufferKey(ih, VK, SC, UD) => (this.Buffer.Push(Format("{{1} {2}}", GetKeyName(Format("vk{:x}sc{:x}", VK, SC)), UD)))
    Start() {
        this.ActiveCount += 1
        SetTimer(this.Stop.Bind(this), -this.Timeout)

        if this.ActiveCount > 1
            return

        this.Buffer := []

        if this.Keybd
            this.InputHook.Start()
        if this.Mouse {
            HotIf this.HotIfIsActive 
            if this.Mouse is String && RegExMatch(this.Mouse, "i)I *(\d+)", &lvl)
                this.SendLevel := Integer(lvl[1])
            opts := this.Mouse is String ? this.Mouse : ("I" this.SendLevel)
            for key in this.MouseButtons {
                if InStr(key, "Wheel")
                    HotKey key, this.BufferMouse.Bind(this), opts
                else {
                    HotKey key, this.BufferMouse.Bind(this,, "Down"), opts
                    HotKey key " Up", this.BufferMouse.Bind(this), opts
                }
            }
            HotIf ; Disable context sensitivity
        }
    }
    Release() {
        if this.IsReleasing
            return []

        sent := [], clickSent := false, this.IsReleasing := 1
        if this.Mouse
            savedCoordMode := A_CoordModeMouse, CoordMode("Mouse", "Screen"), MouseGetPos(&X, &Y)

        ; Theoretically the user can still input keystrokes between ih.Stop() and Send, in which case
        ; they would get interspersed with Send. So try to send all keystrokes, then check if any more 
        ; were added to the buffer and send those as well until the buffer is emptied. 
        PrevSendLevel := A_SendLevel
        SendLevel this.SendLevel - 1
        while this.Buffer.Length {
            key := this.Buffer.RemoveAt(1)
            sent.Push(key)
            if InStr(key, "{Click ")
                clickSent := true
            Send(key)
        }
        SendLevel PrevSendLevel

        if this.Mouse && clickSent {
            MouseMove(X, Y)
            CoordMode("Mouse", savedCoordMode)
        }
        this.IsReleasing := 0
        return sent
    }
    Stop(release := true) {
        if !this.ActiveCount
            return

        sent := release ? this.Release() : []

        if --this.ActiveCount
            return

        if this.Keybd
            this.InputHook.Stop()

        if this.Mouse {
            HotIf this.HotIfIsActive 
            for key in this.MouseButtons
                HotKey key, "Off"
            HotIf ; Disable context sensitivity
        }

        return sent
    }
    GetActiveCount(HotkeyName) => this.ActiveCount
}