global version = 1.30
global appName := "Everyday Helper"
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;===Customised Settings===
#SingleInstance force		;stops complaint message when reloading this file

; Menu, Tray, add, &Menu Name, label_or_shortcut_code
Menu, Tray, add, &About, about
Menu, Tray, add, &Instructions, showInstructions
Menu, Tray, Default, &Instructions 	;Double click icon activates menu item
Menu, Tray, Add, E&xit, ^+Esc
Menu, Tray, NoStandard	;Remove the standard complied hotkey menus: "Exit, Suspend Hotkeys, Pause Script"

; Script Control
^+Esc::ExitApp	;kills application dead when pressing Ctrl+Esc. Note: This line will stop any auto-exec code underneath.
^Esc::	;Reload the script / kill it if there is a problem
	Reload
	Sleep 1000 ;if successful, Reload will close this instance during the Sleep, so the line below will never be reached.
	ExitApp
Return

;===Functions===
PasteValues:	;was my favourite subroutine ;) sac- recommend you put "Sleep, 250" before any GoSub PasteValues if it doesn't work. Think the clipboard can't keep up with lots of SendInput statements
	StringReplace, OutputMe, Clipboard,`r,,A
	SendInput % RegexReplace(OutputMe, "^\s+$")	;trim whitespace including enters (sac- was "^\s+|\s+$")
Return

GetSelection(Trim:=False)	;returns selection
{
	oldClipboard := Clipboard
	Clipboard := ""
	Send ^c
	ClipWait, 2
	If ErrorLevel = 0
		Selection := Clipboard
	Else
		Selection := ""
	If Trim
		Selection = %Selection%
	Clipboard := oldClipboard
	Return Selection
}

PickOneAtRandom(arrayChoice)
{
	Random, index, 1, % arrayChoice.MaxIndex()
	Return % arrayChoice[index]
}

TextFoundInArray(findMe, array)
{
	Loop, Parse, array, `r`n
	{
		if (A_LoopField == findMe)
			return True
	}
	return False
}
;===Main Code===
about:
	Msgbox, 64,,% appName . " by Simon Crouch" . "`n`rversion: " version
Return

showInstructions:
instructions = 
(
EMIS:
Shift+F1		= Find user / Reset password
Ctrl+L		= Paste clipboard (with tabs)

All:
Ctrl+T		= Look up highlighted text in telephone directory
Win+Click		= Ctrl + Printscreen (used for Greenshot)
Ctrl+Esc		= Restart
Ctrl+Shift+Esc 	= Quit
)
	MsgBox, 64, % appName . " - Version " . version, % instructions	
Return

#t::	;look up telephone number in staff address book
	name := GetSelection()
	run chrome.exe http://hr.briscomhealth.org.uk/BCHstaff/fcare.asp?SearchTerm="%name%"
Return

^l::	;Paste Values
	GoSub, PasteValues
Return

;===greenshot/psr alternative===
; Take screenshots (to a folder) when you windows+click. Assumes: Greenshot, Settings, Capture, change Milliseconds to wait before capture = 0
#~LButton::
	SendInput ^{PrintScreen}
Return

#~RButton::
	Sleep 200		;allow time for menus to appear
	SendInput ^{PrintScreen}
Return

;===EMIS Config===
#IfWinActive EMIS Web Health Care System ahk_exe EmisWeb.exe	;General EMIS
	+F1::	;jump to find user search screen
		FindUser()
	Return
#IfWinActive

#IfWinActive Edit user ahk_exe EmisWeb.exe
	+F1::	;reset password
		EditUserResetPassword()
	Return
#IfWinActive

;===EMIS functions===
FindUser(userSearch:="")
{
	IfWinNotActive EMIS Web Health Care System
	{
		Msgbox You are not in the main screen.
		Return
	}
	Send !ezo	;organisation / users config
	WinWait, EMIS Web Health Care System ahk_exe EmisWeb.exe,navBar,5
	If ErrorLevel
	{
		Msgbox Config screen took too long to load.
		Return
	}
	start := A_TickCount
	found := False
	While, (A_TickCount-start <= 3000)	;wait for 5 seconds for screen to laod
	{
		WinGetText, allWinText, A		;find "Users" entry in visible window text (WinText will succeed if exists as part of an entry)
		If (TextFoundInArray("Users", allWinText))
		{
			found := True
			Break
		}
		Sleep 100
	}
	If !(found)
	{
		Msgbox Click Users for me please, then rerun.
		Return
	}
	Send {Alt Down}os{Alt Up}
	WinWaitActive, Find Users
	If userSearch
		Send %userSearch%{Enter}
}

EditUserResetPassword(password:="")
{
	PasswordChoices := ["Password", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
	Msgbox, 4, Reset Password, Do you want to reset this password?
	IfMsgBox No
		Return
	IfWinNotActive Edit user ahk_exe EmisWeb.exe
		Return
	ControlClick, **************, Edit user ahk_exe EmisWeb.exe		;jumps to second password field
	If ErrorLevel
		Return
	If !password
	{
		Random, randNum, 111, 999
		choice := PickOneAtRandom(PasswordChoices)
		password := choice . randNum
	}
	Send {Home}+{End}%password%
	Sleep 150	;required to allow time for EMIS to validate
	Send +{Tab}{Home}+{End}%password%
	Sleep 150	;required to allow time for EMIS to validate
	Control, Check,,User must change password on next sign, Edit user ahk_exe EmisWeb.exe
	Msgbox % "Password set to:`r`n`r`n" . password
}
