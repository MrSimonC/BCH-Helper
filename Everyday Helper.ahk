global version = 2.6
global appName := "Everyday Helper"
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance force		;stops complaint message when reloading this file
#Include notify.ahk
#Include general functions.ahk
#Include instructions.ahk
#Include emis functions.ahk
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
GroupAdd, EMISCodePickOrSelector, Code Selector ahk_exe EmisWeb.exe
GroupAdd, EMISCodePickOrSelector, Code Pick List Properties ahk_exe EmisWeb.exe

; Menu, Tray, add, &Menu Name, label_or_shortcut_code
Menu, Tray, add, &About, about
Menu, Tray, add, &Instructions, showInstructions
Menu, Tray, Default, &Instructions 	;Double click icon activates menu item
Menu, Tray, add
Menu, Tray, add, Open &Derm (8900), openDerm
Menu, Tray, add, Open &Haven (12150), openHaven
Menu, Tray, add, Open &MATS (14232), openMats
Menu, Tray, add
Menu, Tray, add, Open &Train (28159`, 28162), openTrain
Menu, Tray, add, Open &UAT (28167`, 28171), openUAT
Menu, Tray, add, 
Menu, Tray, add, Open Live with &Smartcard (28517), openLiveSmartcard
Menu, Tray, add, Open &Live (28517), openLive
Menu, Tray, add, 
Menu, Tray, Add, E&xit, ^+Esc
Menu, Tray, NoStandard	;Remove the standard complied hotkey menus: "Exit, Suspend Hotkeys, Pause Script"

; Script Control
^+Esc::ExitApp	;kills application dead when pressing Ctrl+Esc. Note: This line will stop any auto-exec code underneath.
^Esc::	;Reload the script / kill it if there is a problem
	Reload
	Sleep 1000 ;if successful, Reload will close this instance during the Sleep, so the line below will never be reached.
	ExitApp
Return

about:
	Msgbox, 64,,% appName . " (BCH Helper) `n`rby Simon Crouch" . "`n`rVersion: " version . "`n`rhttps://github.com/MrSimonC/BCH-Helper"
Return

showInstructions:
	MsgBox, 64, % appName . " - Version " . version, % instructions	
Return

#t::	;look up telephone number in staff address book
	name := GetSelection()
	run chrome.exe http://hr.briscomhealth.org.uk/BCHstaff/fcare.asp?SearchTerm="%name%"
Return

^l::	;Paste Values
	PasteValues()
Return

;greenshot
; Take screenshots (to a folder) when you windows+click. Assumes: Greenshot, Settings, Capture, change Milliseconds to wait before capture = 0
#~LButton::
	SendInput ^{PrintScreen}
Return

#~RButton::
	Sleep 200		;allow time for menus to appear
	SendInput ^{PrintScreen}
Return

;===EMIS Config===
openLive:
:B0*:golive::
	Notify("Live - opening")
	Switcher(28517)
	EmisLogin(28517, A_UserName)
	WinWaitActive, Authentication ahk_exe EmisWeb.exe,,20
	If ErrorLevel
		Return
	SetTitleMatchMode, 2	;text anywhere inside WinTitle
	WinWait, 28517 ahk_exe EmisWeb.exe,,30
Return

openLiveSmartcard:
:B0*:goslive::
	Notify("Live smartcard - opening")
	Switcher(28517)
Return

openHaven:
:B0*:gohaven::
	Notify("Haven - opening")
	Switcher(12150)
	EmisLogin(12150, A_UserName)
	SetTitleMatchMode, 2	;text anywhere inside WinTitle
	WinWait, 12150 ahk_exe EmisWeb.exe,,30
Return

openMats:
:B0*:gomats::
	Notify("MATS - opening")
	Switcher(14232)
	EmisLogin(14232, A_UserName)
	SetTitleMatchMode, 2	;text anywhere inside WinTitle
	WinWait, 14232 ahk_exe EmisWeb.exe,,30
Return

openDerm:
:B0*:goderm::
	Notify("DERM - opening")
	Switcher(8900)
	EmisLogin(8900, A_UserName)
	SetTitleMatchMode, 2	;text anywhere inside WinTitle
	WinWait, 8900 ahk_exe EmisWeb.exe,,30
Return

openTrain:
:B0*:gotrain::
	Notify("TRAIN - opening")
	Switcher(28159)
	EmisLogin(28162, A_UserName)
	SetTitleMatchMode, 2	;text anywhere inside WinTitle
	WinWait, 28162 ahk_exe EmisWeb.exe,,30
Return

openUAT:
:B0*:gouat::
	Notify("UAT - opening")
	Switcher(28167)
	EmisLogin(28171, A_UserName)
	WinWaitActive, Authentication ahk_exe EmisWeb.exe,,10
	If ErrorLevel
		Return
	SetTitleMatchMode, 2	;text anywhere inside WinTitle
	WinWait, 28171 ahk_exe EmisWeb.exe,,30
Return

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

;===(EMIS) Templates===
#IfWinActive ahk_group EMISCodePickOrSelector
	^a::	;Add all codes from selected to the bottom
		Msgbox, 4, Go?, Want down + enter for all these values?, 3
		IfMsgBox Yes
		{
			IfWinActive Code Selector ahk_exe EmisWeb.exe	;as focus can change during execution, ensure the window is in focus
			{
				Send {Enter}
				Loop 500
					Send {Down}{Enter}
			}
		}
	Return

	^+v::	;add in read codes from list in clipboard (copied from Word "Print" of existing template). Optionally select the header & it'll add it as the prompt
			IfWinActive Code Pick List Properties ahk_exe EmisWeb.exe
				ClickReturn(450, 75)	;magnifying glass
			WinWaitActive Code Selector ahk_exe EmisWeb.exe
			entries := StrSplit(Clipboard, "`r`n")
			header =
			Loop % entries.Length()
			{
				IfWinActive Code Selector ahk_exe EmisWeb.exe	;as focus can change during execution, ensure the window is in focus
				{
					code := extractCode(entries[A_Index])
					If(code)
					{
						Send % code
						;Send {Tab}{Enter}{Shift Down}{Tab}{Shift Up}
						Send {Tab}{Enter}		;use this for concepts code selector
					}
					Else
					{
						If(A_Index == 1)
							header := Trim(entries[A_Index])
					}
				}
			}
			;If we have a header, then type it into the prompt box
			WinWaitActive Code Pick List Properties ahk_exe EmisWeb.exe
			If(header != "")
			{
				ClickReturn(130, 210)	;prompt
				Send % header
			}
	Return
	
	^+!v::		;hammer only codes into the code selector window
		entries := StrSplit(Clipboard, "`r`n")
		Loop % entries.Length()
		{
			If (entries[A_Index] == "")	;last entry with a crlf will become "" since it's split above
				Break
			Send % entries[A_Index]
			Send {Down}{Enter}
		}
	Return
#IfWinactive

#IfWinActive Text Pick List Properties ahk_exe EmisWeb.exe
	EnterTextPickList:
	^+v::	;paste multiple text list values from Clipboard (copied from "Print" of existing template)
		entries := StrSplit(Clipboard, "`r`n")
		Loop % entries.Length()
		{
			If (entries[A_Index] == "")	;last entry with a crlf will become "" since it's split above
				Break
			WinWaitActive Text Pick List Properties ahk_exe EmisWeb.exe
			ClickReturn(450, 77) ; +
			WinWaitActive Enter Text List Item ahk_exe EmisWeb.exe
			Send % Trim(entries[A_Index], " `t" . Chr(61551) . Chr(61472))	; Chr(61551)  + Chr(61472) = word square
			Send {Enter}
		}
	Return
	
	^+p::	;paste in the prompt, then the values after
		replaceClipboard := ""
		entries := StrSplit(Clipboard, "`r`n")
		Loop % entries.Length()
		{
			If (entries[A_Index] == "")	;last entry with a crlf will become "" since it's split above
				Break
			If (A_Index == 1)
			{
				prompt := entries[A_Index]
				ClickReturn(175, 209)	;prompt
				Send % prompt
			}
			Else
				replaceClipboard .= entries[A_Index] . "`r`n"
		}
		Clipboard := replaceClipboard
		Sleep 50
		GoSub, EnterTextPickList
	Return
#IfWinActive

#IfWinActive Clinical Code Properties ahk_exe EmisWeb.exe
	^+v::	;paste the single code, found from within the clipboard
		code := extractCode(entries[A_Index])
		If(code)
		{
			ClickReturn(160, 67)	;code
			Send % code
		}
	Return
#IfWinActive

#IfWinActive EMIS Web Health Care System ahk_exe EmisWeb.exe	;General EMIS
	^+s::	;Update component "section". Assumes: correct section key sequence is in clipboard (e.g. "h"="History" section)
		MouseGetPos, currentX, currentY
		If(StrLen(Clipboard)) > 2
		{
			Msgbox Looks like clipboard isn't one or two characters. Stop.
			Return
		}
		Click Right
		Send {Up}{Enter}	;properties
		WinWaitNotActive EMIS Web Health Care System ahk_exe EmisWeb.exe
		ImageClick("images\EMIS\consultation section.png", 215, 10)
		Send % Clipboard
		Send {Enter}
		Sleep 200
		ImageWait("images\EMIS\ok (component properties).png")
		ImageClick("images\EMIS\ok (component properties).png", 10, 10)
		WinWaitActive EMIS Web Health Care System ahk_exe EmisWeb.exe
		MouseMove, %currentX%, %currentY%, 0
	Return
	
	^+c::	;change section to "Create components as children", with Free Text Entry = Title
		Click right
		Send P
		WinWaitActive, Section Properties
		Send ^c{Tab 3}{Space}{Tab}{Down}{Tab}^v
		;Below commented as usually the secion needs changing too after the above action
		;Sleep 500		;time for human to check the screen
		;Send {Enter}
	Return
	
	^+e::		;extract xml files to a location. Non destructive. Assumes save dialogue has correct save location.
		InputBox, loop_number, Loop Number, Number of times to loop:,,300,130
		If ErrorLevel
			Return
		Loop, %loop_number%
		{
			Send !lw
			WinWaitActive Save As
			Send !s	;save
			Sleep 100
			IfWinActive Confirm Save As
			{
				Send !n	;no
				WinWaitActive Save As
				Send {Esc}	;Cancel
				Notify("File exists already.", "Stop.")
				Send +{Tab}{Down}{Tab 2}{Down}{Up} ;move to next folder
				Return
			}
			WinWaitActive EMIS Web Health Care System ahk_exe EmisWeb.exe
			Send {Down}
		}
	Return
	
	^+p::	;save template as word print. Assumes save dialogue has correct save location.
		Send {AppsKey}p
		Sleep 200
		templateName := GetSelection()
		Send {Esc}
		WinWaitActive EMIS Web Health Care System ahk_exe EmisWeb.exe	;General EMIS
		Send !lpt
		WinWaitActive Template Preview
		Sleep 500		;give time to load page
		Send {Tab 2}{Space}
		WinWaitActive Save As
		Send % templateName
		Send {Enter}
		WinWaitActive EMIS Web Health Care System ahk_exe EmisWeb.exe
		Send {Down}
	Return
#IfWinActive

#IfWinActive Term Properties ahk_exe EmisWeb.exe
	^+v::	;Care Summary view, Add Code, Term Properties - add code lists from Clipboard (use "Clinical Views.xlsx"). Assumes: you've copied the list of codes, top down, from the xls sheet
		KeyWait, Shift
		KeyWait, Control
		WinWaitActive Term Properties ahk_exe EmisWeb.exe
		ClickReturn(430, 100)		;magnifying glass
		WinWaitActive Code Selector ahk_exe EmisWeb.exe
		entries := StrSplit(Clipboard, "`r`n")
		Loop % entries.Length()
		{
			If (entries[A_Index] == "")	;last entry with a crlf will become "" since it's split above
				Break
			IfWinActive Code Selector ahk_exe EmisWeb.exe	;as focus can change during execution, ensure the window is in focus
			{
				Send % extractCode(entries[A_Index])
				Send {Tab}{Enter}{Shift Down}{Tab}{Shift Up}
			}
		}
		Send {Shift Down}{Tab 3}{Shift Up}{Space}	; cancel
		; fully automate by copying header from excel, then pressing ok and moving to the bottom
		WinActivate ahk_exe EXCEL.EXE
		Sleep 300
		Send {Left}^c
		Sleep 300
		WinActivate Term Properties ahk_exe EmisWeb.exe
		WinWaitActive Term Properties ahk_exe EmisWeb.exe
		ClickReturn(137, 72)	;Display Term
		Send ^v!o	;Display Term + OK
		WinWaitActive Edit Code List Section ahk_exe EmisWeb.exe
		Send ^{End}	;go to bottom
		Send {Alt Down}{Alt Up}{Right}{Enter}	;Add Code
	Return
#IfWinActive

#IfWinActive ahk_class OpusApp		;Word
	^+c::	;copies and cleans the word output from EMIS template "print" to a pastable format (removes square tick boxes)
		Send ^c
		Sleep 100
		output := ""
		entries := StrSplit(Clipboard, "`r`n")
		Loop % entries.Length()
		{
			If (entries[A_Index] == "")	;last entry with a crlf will become "" since it's split above
				Break
			output .= cleanWordLine(entries[A_Index]) . "`r`n"
		}
		Clipboard := output
		Notify("Copy && Clean", "Complete", 2)
	Return
#IfWinActive

::gogetcon::	;save consultation (in full) as a word document (to path: TMP, TEMP, or USERPROFILE). Will stop if no patient already in context. Saves document save path to clipboard
	savePath := A_Temp . "\"
	WinActivate ahk_exe EmisWeb.exe
	Send !ecc		;consultations
	Sleep 1000
	WinWaitActive, Patient Find ahk_exe EmisWeb.exe,,1
	If !ErrorLevel
		Return
	Send !cpi	;Print Fully Summary with Attachments
	WinWaitActive Full Patient Summary ahk_exe EmisWeb.exe
	Sleep 2000		;left word load
	Click 250, 250	;click into Word, else F12 won't work
	Send {F12}	;save as
	WinWaitActive Save As
	filename := GetSelection()
	Send {Home}
	Send % savePath
	Send {Enter}
	Sleep 2000	;allow time to save
	WinWaitActive ahk_exe WINWORD.EXE
	Send !{f4}
	Clipboard := savePath . filename
Return

#IfWinActive Drop-Down Form Field Options ahk_exe WINWORD.EXE
	::goadd::	;Add clipboard entries to EMIS Document: Legacy Drop-Forwn Form Field
		entries := StrSplit(Clipboard, "`r`n")
		Loop % entries.MaxIndex()
		{
			if(entries[A_Index] == "")		;Skip blanks
			Continue
			Send % entries[A_Index]
			Send !a
		}
	Return
#IfWinActive

::gosetcon::	;Attach document to user. Assumes path to document is in clipboard
	IfNotExist, %Clipboard%
	{
		Msgbox Clipboard doesn't contain a valid path to a file.
		Return
	}
	WinActivate ahk_exe EmisWeb.exe
	Send !ecc		;consultations
	Sleep 1000
	WinWaitActive, Patient Find ahk_exe EmisWeb.exe,,1
	If !ErrorLevel
		Return
	Send !cadz		;Attach document
	WinWaitActive Open ahk_exe EmisWeb.exe
	Send % Clipboard
	Send {Enter}
	WinWaitActive Attach Document ahk_exe EmisWeb.exe
	Sleep 2000	;let it load
	Send {Tab 2}{Enter}
	CodeSelectorAddSingleCode("9l5")		;code: "patient record merged"
	Send {Tab 7}	;Document title
Return
