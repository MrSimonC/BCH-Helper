global version = 1.23
global appName := "BCH Helper"
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;===Customised Settings===
#SingleInstance force		;stops complaint message when reloading this file

; Menu
; Menu, Tray, add, &Menu Name, label_or_shortcut_code
Menu, Tray, add, &About, about
Menu, Tray, add, &Instructions, showInstructions
; Menu, Tray, Default, &Menu Name 	;Double click icon activates menu item
Menu, Tray, Add, E&xit, ^+Esc
Menu, Tray, NoStandard	;Remove the standard complied hotkey menus: "Exit, Suspend Hotkeys, Pause Script"

; Script Control
^+Esc::ExitApp	;kills application dead when pressing Ctrl+Esc. Note: This line will stop any auto-exec code underneath.
^Esc::	;Reload the script / kill it if there is a problem
Reload
Sleep 1000 ;if successful, Reload will close this instance during the Sleep, so the line below will never be reached.
ExitApp
Return

;=====================

about:
Msgbox, 64,,% appName . " by Simon Crouch" . "`n`rversion: " version
Return

showInstructions:
instructions = 
(
Ctrl T	= Telephone: look up highlighted text in the directory
Shift F1	= EMIS: Find user
Shift F1	= EMIS Edit user: Reset user password
Win+Click	= Presses Ctrl + Printscreen (used for Greenshot)
)
MsgBox, 64,, % instructions, 
Return

#t::	;look up telephone number in staff address book
name := GetSelection()
run chrome.exe http://hr.briscomhealth.org.uk/BCHstaff/fcare.asp?SearchTerm="%name%"
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

===EMIS functions===
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
	Sleep 500	;screen needs more time to load
	WinGetText, allWinText, A		;find "Users" entry in visible window text (WinText will succeed if exists as part of an entry)
	If !(TextFoundInArray("Users", allWinText))	
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
	Sleep 150	;required to allow time for EMIS to validate. test push - remove this comment
	Control, Check,,User must change password on next sign, Edit user ahk_exe EmisWeb.exe
	Msgbox % "Password set to:`r`n`r`n" . password
}

;===Functions===
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