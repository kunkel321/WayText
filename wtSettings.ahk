#SingleInstance Force
#Requires AutoHotkey v2+
;===============================================================================
; This is the wtSettings code.	By Kunkel321 5-2-2024							
; These are the settings for the WayText application. 				
; The script file is not a "stand alone" tool -- the folder and ini files are needed. 
; Setting values are added via this Gui, then saved to the ini, then read by WayText.
;===============================================================================

TraySetIcon("shell32.dll","81")   ; I think this is the old "disc defrag" icon?

settingsFile := "wtFiles\Settings.ini"
If not FileExist(settingsFile) 
{	Msgbox "We must exit because the 'Settings.ini' file was not found at location:`n`n" A_ScriptDir settingsFile, "WAYTEXT SETTINGS", 48
	ExitApp
}

;==========Read from Settings file========
varHotkey := IniRead(settingsFile, "MainSettings", "MyHotkey")
varVisualMode := IniRead(settingsFile, "MainSettings", "VisualSendMode", "0")
useBuffer := IniRead(settingsFile, "MainSettings", "UseInputBuffer", "1")
postPasteDelay := IniRead(settingsFile, "MainSettings", "PostPasteDelay", "200")
varHotKey := IniRead(settingsFile, "MainSettings", "MyHotkey", "!+w")
varFaveEditor := IniRead(settingsFile, "MainSettings", "FaveEditor")
varTargetWins := IniRead(settingsFile, "MainSettings", "PreferredTargetWins")
varForcePref := IniRead(settingsFile, "MainSettings", "ForceUseOfPreferred")
varColor := IniRead(settingsFile, "MainSettings", "GUIColor")
varFontColor := IniRead(settingsFile, "MainSettings", "FontColor")
varListColor := IniRead(settingsFile, "MainSettings", "ListColor")
varWColor := IniRead(settingsFile, "MainSettings", "GUIwarnColor")
varDefaultMenuColor := IniRead(settingsFile, "MainSettings", "MenuColorIsDefault", "1")
varDebug := IniRead(settingsFile, "MainSettings", "Debug")
varLastTab := IniRead(settingsFile, "MainSettings", "LastTab")
varMon := IniRead(settingsFile, "schedule", "Monday")
varTue := IniRead(settingsFile, "schedule", "Tuesday")
varWed := IniRead(settingsFile, "schedule", "Wednesday")
varThu := IniRead(settingsFile, "schedule", "Thursday")
varFri := IniRead(settingsFile, "schedule", "Friday")
; varSat := IniRead(settingsFile, "schedule", "Saturday")
; varSun := IniRead(settingsFile, "schedule", "Sunday")
varMemoryMinutes := IniRead(settingsFile, "MainSettings", "MemoryMinutes", "60")
varDefName  := IniRead(settingsFile, "MainSettings", "defaultName", "Student")
varMyColors := IniRead(settingsFile, "ColorsList")

;======Set color names and hex codes from above ini values=====================
varColorName := subStr(varColor, 1, -8) ; left part of var is color name.
varColorCode := subStr(varColor, -6) ; right part of var is color hex.
varListColorName := subStr(varListColor, 1, -8)
varListColorCode := subStr(varListColor, -6)
varFontColorName := subStr(varFontColor, 1, -8)
varFontColorCode := subStr(varFontColor, -6)
varWColorName := subStr(varWColor, 1, -8)
varWColorCode := subStr(varWColor, -6)

;======Launched from WayText app?=========================
Try WinGetPos(&wtsX, &wtsY, &wtsW, &wtsH, "A")   	; "A" to get the active window's position.
If isSet(wtsW) 
{	if WinActive("WayText")
	{	;msgbox 'wts active'
		wtsX := wtsX + wtsW ;If launched from WT form, open right next to it.
	}
	else
	{	wtsX := wtsX + (wtsW * 0.05) ; Use these with GUI Show, below.
		wtsY := wtsY + (wtsH * 0.2)
	}
}
If not isSet(wtsW)
{	wtsX := A_ScreenWidth / 4 ; Near top/left of screen.
	wtsY := A_ScreenHeight / 4
}

;=======Build the GUI Form=========
wts := Gui()
wts.Opt("-MinimizeBox +alwaysOnTop")
wts.SetFont("c" varFontColorCode " s11")
wts.BackColor := varColorCode

myTabs := ["Activation","Send Mode","Editor","Targets","Colors","Schedule","Pre-Fill","Tips"]
; If myTabs array changes, please copy/paste to "menu" code in WayText script.
if (A_Args.Length > 0) ; Check if a command line argument is present and set the default tab
    varLastTab := A_Args[1]
; AltSubmit so we can save the Tab number and open with same (A_Args overrides this).
Tab := wts.Add("Tab3", "AltSubmit vCurrTab Choose" . varLastTab, myTabs)

; The 8 tabs...
;###############################################################################################
;###############################################################################################
;###############################################################################################
Tab.UseTab(1) ;######## ACTIVATION #####################################

wts.Add("Text", "Wrap  w340", "The current hotkey is indicated below. Change if desired.")
wHotkey := wts.Add("Hotkey", , varHotkey)
wts.Add("Text", "Wrap  w340", "Tip: The Win key (#) modifier cannot be added here.  It can by manually entered into the Settings.ini file, but if you use it, the above box will show `"None`".")

; start with windows? 
if FileExist(A_Startup "\WayText.lnk")
{	StartCheck := "Checked"
	StartTxt := "Unheck to no longer start with Windows."
}
else
{	StartCheck := ""
	StartTxt := "Check to start with Windows."
}
(StartChkBox := wts.Add("CheckBox", StartCheck, StartTxt)).OnEvent("Click", StartUp)

StartUp(*)
{	If (StartChkBox.Value = 1)
	{	FileCreateShortcut(A_WorkingDir "\WayText.exe", A_Startup "\WayText.lnk")
		MsgBox("WayText will auto start with Windows.", "WayText Settings", 4096)
	}
	Else If (StartChkBox.Value = 0)
	{	FileDelete(A_Startup "\WayText.lnk")
		MsgBox("WayText will NO LONGER auto start with Windows.", "WayText Settings", 4096)
	}
}

;###############################################################################################
Tab.UseTab(2) ;######### SEND MODE ####################################

if (varVisualMode = 1)
{	ModeVisual := "checked"
	ModeSafer := ""
}
else
{	ModeVisual := ""
	ModeSafer := "checked"
}

wts.add("groupbox", "x30 y80 h90 w340", "Text Type Send Modes")
wts.Add("Radio", "yp+30 xp+10 " . ModeSafer, "Safer and Faster, but works in Background.")
wtsRadioVisMode := wts.Add("Radio", "vVisMode " . ModeVisual, "More Visual, but prone to errors.")

wts.Add("Text", "x30 yp+40 Wrap", "Click for Tips about Send Modes.").OnEvent("Click", TipsSendMode)
TipsSendMode(*)
{	SendModeTips := "
	(
		====Text Type Send Modes:====
		AutoHotkey's `"SendInput`" is Safer and Faster but works in Background.  You might not see the fields in the webform getting updated, but Chrome will update the whole page after the text is sent.

		`"SendEvent`" is More Visual, but prone to errors. The keys are send as separate events, so Chrome refreshes after each character.  

		When using SendInput, AutoHotkey tries to buffer any keypresses that occur during the auto-typing of the boilerplate text.  This is an imperfect process in Windows though.  Descolada's InputBuffer Class addresses this and `"captures`" extraneous keypresses during auto-typing, then places the captured text.  Learn more on the forum thread at the AutoHotkey Forums. 
	)"
	Msgbox(SendModeTips,,4096+64)
}

wInputBuffer := wts.add("checkbox", "x30 yp+40 Checked" useBuffer, "Use Descolada's InputBuffer Class`n(recommended).")
wts.setFont("underline")
wts.Add("Text", "cBlue", "Click for InputBuffer forum thread").OnEvent("click"
, (*) => Run("https://www.autohotkey.com/boards/viewtopic.php?f=83&t=122865"))
wts.setFont("norm")

wtsPostPasteDelay := wts.Add("Edit", "w60 yp+50 Background" varListColorCode, )
wts.Add("UpDown", "Range0-1000", PostPasteDelay)
wts.Add("Text", "Wrap X+15 yp-5", "Post-Paste Delay (milliseconds). Setting`nit to zero essentially disables the delay.")

;###############################################################################################
Tab.UseTab(3) ;####### EDITOR #########################################
wts.Add("text", "Wrap w340", "The ini files can be opened from the WayText SysTray menu.  Choose ini editor to use.")

If (StrLen(varFaveEditor)>48) ; If the path is too long, it makes the GUI too wide.
	EditorLBL := SubStr(varFaveEditor, 1, 3) "...." SubStr(varFaveEditor, -41)
Else
	EditorLBL := varFaveEditor

wts.Add("text", , EditorLBL)
(buttChooseEditor := wts.add("button", , "Choose Different Editor")).OnEvent("Click", ChooseEditorFunc)
(buttUseSciTE := wts.add("button", , "Use Embeded SciTE")).OnEvent("Click", UseSciteFunc)
ChooseEditorFunc(*)
{	SelectedEditor := FileSelect(3, "", "Choose favorite editor", "Application (*.exe)")
	if (SelectedEditor = "")
		MsgBox("No app was selected.", "WayText Settings", 4096)
	else
		IniWrite(SelectedEditor, settingsFile, "MainSettings", "FaveEditor") ; IniWrite happens here instread of bottom of code. 
	Reload()
}
UseSciteFunc(*)
{	if FileExist("SciTE\SciTE.exe")
		IniWrite(A_ScriptDir "\SciTE\SciTE.exe", settingsFile, "MainSettings", "FaveEditor") ; IniWrite happens here instread of bottom of code. 
	else
		MsgBox("Scintilla Text Editor (SciTE) not found bundled with WayText.", "WayText Settings", 4096)
	Reload()
}
wts.setFont("underline")
wts.Add('text', 'cBlue x+20', 'Info about SciTE').OnEvent('click', (*) => MsgBox("From SciTE help file: `"SciTE distribution designed for AutoHotkey - made by fincs - Original SciTE made by Neil Hodgson.`"`n`nYou can get the Scintilla Text Editor for AHK (SciTE4AHK) from the AutoHotkey website.",,4096+64))
wts.setFont("norm")

wts.Add("Text", "Wrap w340 x30 y224", "The Definitions and the Offices configuation (ini) files can also be opened from here. Click to open.")
wts.Add("Button", "section", "List of Offices").OnEvent("Click", (*) => Run("wtFiles\OfficeList.ini"))
; wts.Add("Button", "xp+24 yp+80 section", "List of Offices").OnEvent("Click", (*) => Run("wtFiles\OfficeList.ini"))
wts.Add("Button", , "List of Boilerplate (Definition Library) Entries").OnEvent("Click", (*) => Run("wtFiles\TextDefinitions.ini"))

;###############################################################################################
Tab.UseTab(4) ;####### TARGETS ########################################
wts.Add("Text", "Wrap  w340", "`"Preferred`" target windows are those that you intend to send text input to.  Enter a list of preferred windows, separated by comma space to help prevent this. RegEx partial-match rules used to identify windows by their titles. Note: `".*`" (dot asterisk) matches everything.")
varTargetWins := StrReplace(varTargetWins, "|", ", ") ; Comment-out this, if you prefer the|list|like|this.

wtsTargetWins := wts.Add("Edit", "+wrap  w340 h42 cBlack  Background" varListColorCode, varTargetWins)
wts.add("text", "+wrap  w340", "Check below to only type into preferred windows. WayText uses a popup MessageBox if accidentally triggered in wrong window. Tip: Holding <Shift> forces use of MsgBox regardless of window. This is `"Debug Mode`".")
wtsChkForcePref := wts.Add("CheckBox", "xs+10 vChkForcePref checked" . varForcePref, "Use MsgBox if target window is non-preferred.")

;###############################################################################################
Tab.UseTab(5)  ;####### COLORS #########################################
wColorTips := wts.Add("Text", "Wrap  w340", "Click for Tips about these colors.")
wColorTips.OnEvent('Click', colorTipsMsg)
colorTipsMsg(*)
{
	vcolorTipsMsg := "
	(
		It's okay to enter a custom color, but it must be entered as a hex value.  Neither RBG, nor HTML color names, are recognized.`n`nIf desired, for a more permanent way to add custom colors, you can add items to the long list of colors.  To do so, open the Settings.ini file and go to the bottom.  Make sure to add your new color using the same format:
		`n`n--------------------`nColor Name, hexcode`n--------------------`n`n
		Explanation of the checkbox:`n
		`"Use Windows default for menu background color`" This is for when a Dark GUI color, and light Font are used for the theme.  The font used in the Windows System Tray Context Menu cannot be colorized.  Therefore, you'll always want a light backcolor.  Setting it to the Windows default will ensure that you can see the text. 

	)"
	MsgBox(vcolorTipsMsg,,4096+64)
}

arrMyColors := strSplit(varMyColors, "`n")

ColorChoose := 1 ; needed this as default to prevent error.
For idx, color in arrMyColors {
	if color = varColor
	{	ColorChoose := idx ; Get position of current color for 'Choose' option.
		Break
	}
}
wCmbColor := wts.Add("ComboBox", "w160 cDefault  Section Choose" . ColorChoose " Background" varListColorCode, arrMyColors)
wCmbColor.OnEvent("Change", guiColorChange)
guiColorChange(*)
{	wts.BackColor := subStr(wCmbColor.Text, -6)
	wtstextTxtForm.Text := "Form: " subStr(wCmbColor.Text, 1, -8)
}

ListColorChoose := 1 ; default to first item, which is white. 
For idx, color in arrMyColors {
	if color = varListColor
	{	ListColorChoose := idx
		Break
	}
}
wCmbListColor := wts.Add("ComboBox", "w160 cDefault  Choose" . ListColorChoose " Background" varListColorCode, arrMyColors)
wCmbListColor.OnEvent("Change", wListColUpdate)
wListColUpdate(*)
{	wProgList.Opt("c" subStr(wCmbListColor.Text, -6))
	wtstextTxtList.Text := "List Boxes: " . subStr(wCmbListColor.Text, 1, -8)
	wProgListLbl.Text := "Sample of List color: " subStr(wCmbListColor.Text, 1, -8) "."
}

FontColorChoose := arrMyColors.Length ; default to last item, which is black. 
For idx, color in arrMyColors {
	if color = varFontColor
	{	FontColorChoose := idx
		Break
	}
}
wCmbFontColor := wts.Add("ComboBox", "w160 cDefault  Choose" . FontColorChoose " Background" varListColorCode, arrMyColors)
wCmbFontColor.OnEvent("Change", updateFontColor)
updateFontColor(GuiCtrlObj, *)
{	wtstextTxtFont.Text := "Font: " subStr(wCmbFontColor.Text, 1, -8)
	For Ctrl in [wColorTips, wtstextTxtForm, wtstextTxtList, wtstextTxtFont, wtstextTxtWarning, wProgListLbl, wProgWListLbl]
            Ctrl.Opt("c" subStr(wCmbFontColor.Text, -6)) ; For Loop colorizes only the text on this Tab. 
}

WColorChoose := 1 ; default to first item, which is white. 
For idx, color in arrMyColors {
	if color =  varWColor
	{	WColorChoose := idx
		Break
	}
}		
wCmbWarnColor := wts.Add("ComboBox", "w160 cDefault  Choose" . WColorChoose " Background" varListColorCode, arrMyColors)
wCmbWarnColor.OnEvent("Change", wListWarnColUpdate)
wListWarnColUpdate(*)
{	wProgWListColor.Opt("c" subStr(wCmbWarnColor.Text, -6))
	wtstextTxtWarning.Text := "Warning: " subStr(wCmbWarnColor.Text, 1, -8)
	wProgWListLbl.Text := "Sample of List color: " subStr(wCmbWarnColor.Text, 1, -8) "."
}

wtsDefaultMenuColor := wts.Add("Checkbox", " Checked" varDefaultMenuColor, "Use Windows default for menu background color.")

wtstextTxtForm := wts.Add("text", "ys x200 w180 Section", "Form: " .  SubStr(varColorName, 1, 13)) ; Only enough room on form for first 13 letters. 
wtstextTxtList := wts.Add("text", "w180", "List Boxes: " . SubStr(varListColorName, 1, 13))
wtstextTxtFont := wts.Add("text", "w180", "Font: " . SubStr(varFontColorName, 1, 13))
wtstextTxtWarning := wts.Add("text", "w180", "Warning: " . SubStr(varWColorName, 1, 13))

wProgList := wts.Add("Progress", "vwProgList x28  yp+52 w330 h50 c" . varListColorCode, "100")
wProgListLbl := wts.Add("text", "xp+14 yp+7 w323 +BackgroundTrans", "Sample of List background color:`n" . varListColorName . ".")

wProgWListColor := wts.Add("Progress", "vwProgWListColor xp-14 yp+50 w330 h50 c" . varWColorCode, "100")
wProgWListLbl := wts.Add("text", "xp+14 yp+7 w323 +BackgroundTrans", "Sample of Form color when activated from a`nnon-preferred window: " . varWColorName . ".")

;###############################################################################################
Tab.UseTab(6) ;###### SCHEDULE #########################################
wts.Add("Text", "Wrap  w340", "Choose below which office should be the default selection for each day of the week.")
wts.Add("Text", "section", "Monday")
wts.Add("Text", , "Tuesday")
wts.Add("Text", , "Wednesday")
wts.Add("Text", , "Thursday")
wts.Add("Text", , "Friday")
; wts.Add("Text", , "Saturday") ; Un-comment-out and use if desired.
; wts.Add("Text", , "Sunday")

AllOffices := IniRead("wtFiles\OfficeList.ini") ; Gets `n-delimited list from ini file.
arrSchedule := strSplit(AllOffices, "`n")
varMonBld := wts.Add("DropDownList", "ys section w130 AltSubmit vvarMonBld Choose" . varMon, arrSchedule)
varTueBld := wts.Add("DropDownList", "w130 vvarTueBld AltSubmit Choose" . varTue, arrSchedule)
varWedBld := wts.Add("DropDownList", "w130 vvarWedBld AltSubmit Choose" . varWed, arrSchedule)
varThuBld := wts.Add("DropDownList", "w130 vvarThuBld AltSubmit Choose" . varThu, arrSchedule)
varFriBld := wts.Add("DropDownList", "w130 vvarFriBld AltSubmit Choose" . varFri, arrSchedule)
; varSatBld := wts.Add("DropDownList", "w130 vvarSatBld AltSubmit Choose" . varSat, arrSchedule)
; varSunBld := wts.Add("DropDownList", "w130 vvarSunBld AltSubmit Choose" . varSun, arrSchedule)

;###############################################################################################
Tab.UseTab(7) ;###### PREFILL #########################################
wts.Add("Text", "Wrap  w340", "WayText will remember the Name, Gender, Office, and Boilerplate Text Definition that was used with the last text entry. This memory will override the office chosen in the Schedule tab.  Last used information should be prefilled if the last use was within this many minutes. Choose from 1-600. If WayText has been unused for more minutes than this, the default name, gender, Office, Definition values will be used.")

wtsMemoryEdit := wts.Add("Edit", "vMemoryEdit w60 cBlack Background" varListColorCode)
wtsUpDownMemoryMinutes := wts.Add("UpDown", "vMemoryMinutes Range0-600", varMemoryMinutes)
wts.Add("Text", "Wrap w340", "Setting it to zero essentially disables the memory.")

wts.Add("Text", "Wrap w340", "Assign default text to appear in `"Name`" box.`n(Student, Patient, etc.)")
wtsDefaultName := wts.Add("Edit", "w100 Background" varListColorCode, varDefName)

;##############################################################################################
Tab.UseTab(8) ;###### TIPS #########################################
wts.setFont("s13")
wts.Add("Text", "w340", "     WAYTEXT SETTINGS")
wts.setFont("s11")
wts.Add("Picture", "w240 h120", A_ScriptDir "\icons\wtSettingsLogo.ico")

tipTabText := "
(
	[Apply] >> Save Settings.  If launched from WayText, reload and open WayText.
	[Apply Then Close] >> Save Settings and Close This, If launched from WayText, reload, but don't open WayText.
)"
wts.Add("Text", "Wrap w340", tipTabText)
wts.Add("Text", "w340", "Click for more tips.").OnEvent("click", tipTabMore)
tipTabMore(*)
{ tipTabMsg := "
	(
		These are the settings for the WayText application. 

		They save the settings to the Settings.ini file, then WayText reads from the ini file each time it is restated.  As such, WayText.exe gets restarted when setting changes are made here, using the following rules: 

		If these settings were launched from the WayText main form, then Pressing the [Apply] button at the bottom will save the settings, reload WayText, then open its form.  

		Pressing [Apply Then Close] saves and reloads, then closes the Settings form, but does not open the WayText form. 

		If Settings is not launched via WayText, then [Apply] only saves the settings to the ini file, and [Apply Then Close] saves to the ini file, reloads WayText without opening its form, and closes the settings form.

		Pressing [Cancel] exits the Setting Form and does NOT save the settings to the ini file. 
	)"
	Msgbox(tipTabMsg,,4096+64)
}

;###############################################################################################
;###############################################################################################
;###############################################################################################
Tab.UseTab() ;###### END OF TABS #########################################

;======Add bottom part of Gui==========
wts.Add("Button", "w106 h30 Section", "Apply").OnEvent("click", (*) => IniWrites("Apply"))
wts.Add("Button", "ys w136 h30", "Apply Then Close").OnEvent("click", (*) => IniWrites("ApplyNClose"))
wtsButtonCancel := wts.Add("Button", "ys w106 h30", "Cancel").OnEvent('click', exitSettings)
wts.Title := "wtSettings"
;=======Display/Open the main form=======
wts.Show("x" . wtsX . " y" . wtsY)

; [Cancel] button calls this function. 
exitSettings(*)
{	wts.Destroy
}

; [Apply] and [Apply then close] buttons both call this function.
IniWrites(closeOrNot)
{	; Save the above entered settings to the Settings.ini file so that WayText can read them.
	; "LastTab is the number for open tab, so we can restore it.  Choosing a Tab via the systray
	; sub menu will override this however. 
	IniWrite(Tab.Value, settingsFile, "MainSettings", "LastTab")
		;======= Tab 1 HotKey ===================
	IniWrite(wHotkey.Value, settingsFile, "MainSettings", "MyHotkey")
		;======= Tab 2 Send text mode =====================
	;VisMode -= 1
	IniWrite(wtsRadioVisMode.Value, settingsFile, "MainSettings", "VisualSendMode")
	IniWrite(wInputBuffer.Value, settingsFile, "MainSettings", "UseInputBuffer")
	IniWrite(wtsPostPasteDelay.Value, settingsFile, "MainSettings", "PostPasteDelay")
		;========  Tab 3 Fave Editor ======================
	; IniWrites happen on demand, in the Fave Editor tab.
		;======== Tab 4 Fave Windows ======================
	global varTargetWins := StrReplace(wtsTargetWins.Text, ", ", "|") ; Put back into regex format for ini file. 
	IniWrite(varTargetWins, settingsFile, "MainSettings", "PreferredTargetWins")
	IniWrite(wtsChkForcePref.Value, settingsFile, "MainSettings", "ForceUseOfPreferred")
		;======= Tab 5 Colors =============================
	IniWrite(wCmbColor.Text, settingsFile, "MainSettings", "GUIColor")
	IniWrite(wCmbListColor.Text, settingsFile, "MainSettings", "ListColor")
	IniWrite(wCmbFontColor.Text, settingsFile, "MainSettings", "FontColor")
	IniWrite(wCmbWarnColor.Text, settingsFile, "MainSettings", "GUIwarnColor")
	IniWrite(wtsDefaultMenuColor.Value, settingsFile, "MainSettings", "MenuColorIsDefault")
		;======= Tab 6 Schedule ===========================
	IniWrite(varMonBld.Value, settingsFile, "schedule", "Monday")
	IniWrite(varTueBld.Value, settingsFile, "schedule", "Tuesday")
	IniWrite(varWedBld.Value, settingsFile, "schedule", "Wednesday")
	IniWrite(varThuBld.Value, settingsFile, "schedule", "Thursday")
	IniWrite(varFriBld.Value, settingsFile, "schedule", "Friday")
	; IniWrite(varSatBld.Value, settingsFile, "schedule", "Saturday")
	; IniWrite(varSunBld.Value, settingsFile, "schedule", "Sunday")
		;======= Tab 7 Prefill ===========================
	IniWrite(wtsMemoryEdit.Value, settingsFile, "MainSettings", "MemoryMinutes")
	IniWrite(wtsDefaultName.Value, settingsFile, "MainSettings", "defaultName")
		;======= Tab 8 Tips ==============================
	; No settings to save for tab 8...

	; Reload WayText in RAM (and press hotkey, when appropriate.)
	Run A_ScriptDir "\WayText.exe"

	If WinActive("WayText Application") and closeOrNot = "Apply"
	{	While not processExist("WayText.exe")
			Sleep 50 ; Wait for it to be running again before trying the hotkey. 
		sleep 800 ; Wait more, just to be safe.
		Send '"' varHotKey '"'
	}
	wts.Destroy() ; Destroy settings form.
	
}
