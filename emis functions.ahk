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

EditUserResetPassword(password:="", askToEmail:=True)
{
	PasswordChoices := ["Password", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
	;Msgbox, 4, Reset Password, Do you want to reset this password?
	;IfMsgBox No
	;	Return
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
	Send {Shift Down}{Tab 2}{Shift Up}	;email
	emailAddress := GetSelection()
	Control, Check,,User must change password on next sign, Edit user ahk_exe EmisWeb.exe
	Msgbox % "Password set to:`r`n`r`n" . password
	If askToEmail
	{
		MsgBox, 4, Email end user?, Send email to %emailAddress%?
		IfMsgBox, Yes
			SendPasswordResetEmail(emailAddress, password)
	}
}

Switcher(switcher_id)
{
	emis_url = C:\Users\Public\Desktop\EMIS Web.url
	IfNotExist % emis_url
	{
		Msgbox % emis_url . " doesn't exist. Can you add/install it?"
		Exit
	}
	if org_id == ""
		org_id = switcher_id
	If !WinExist("EMIS Configuration Switcher")
	{
		Run % emis_url
		Sleep 1000
	}
	WinActivate ahk_exe SDS.Client.ConfigurationSwitcher.exe
	WinWaitActive ahk_exe SDS.Client.ConfigurationSwitcher.exe
	ControlSetText Edit1,%switcher_id%, ahk_exe SDS.Client.ConfigurationSwitcher.exe
	ControlSend, Edit1, {Enter}, ahk_exe SDS.Client.ConfigurationSwitcher.exe
	WinWait, Authentication ahk_exe EmisWeb.exe,,5		;login screen. smartcard doesn't require this screen, so wait for 5 then kill switcher anyway
	WinClose ahk_exe SDS.Client.ConfigurationSwitcher.exe
}

EmisLogin(org_id, username, pass="")
{
	WinWait Authentication ahk_exe EmisWeb.exe
	WinActivate Authentication ahk_exe EmisWeb.exe
	ClickReturn(350, 140)	;username
	Send %username%{Tab}%pass%{Tab}%org_id%
	Send {Shift Down}{Tab}{Shift Up}
}

extractCode(string)	;returns the EMIS/SNOMED code from a string
{
	RegExMatch(string, "(\b[0-9]+[a-zA-Z0-9\-]+|[a-zA-Z0-9\-]+[0-9]+\b|\bEMISREQ\|[0-9]+[a-zA-Z0-9\-]+\.*)", code)
	Return % code
}

cleanWordLine(line)		;takes one line from EMIS template Word Print, outputs: Code{Tab}Description
{
	line := StrReplace(line, A_Tab)
	line := StrReplace(line, " - ", A_Tab, , 1)	;replace only first " - "
	line := StrReplace(line, chr(61551)) ;unicode Wingdings box
	line := StrReplace(line, chr(61472)) ;unicode Wingdings box
	Return % line
}

CodeSelectorAddSingleCode(code)
{
	WinWaitActive Code Selector ahk_exe EmisWeb.exe
	Send % code
	Send {Down}{Enter}
	WinWaitNotActive Code Selector ahk_exe EmisWeb.exe
}

GoToConsultations()
{
	WinActivate ahk_exe EmisWeb.exe
	Send !ecc		;consultations
	WinWaitActive, Patient Find ahk_exe EmisWeb.exe,,1
	If !ErrorLevel
		Return False
	Return True
}

AddLetter(letterName := "")
{
	Send !cadl	;attach letter
	WinWaitActive New Patient Letter ahk_exe EmisWeb.exe
	Click 324, 314	;magnifying glass
	WinWaitActive Find Document Templates ahk_exe EmisWeb.exe
	Send % letterName
	Send {Enter}{Down}{Enter}
	WinWaitNotActive Find Document Templates ahk_exe EmisWeb.exe
}

SendPasswordResetEmail(to, newPassword)
{
	FileRead, htmlBody, assets/passwordResetBody.htm
	StringReplace, htmlBody, htmlBody, $pass, % newPassword
	Email(False, to, "Your EMIS password has been reset.","", htmlBody,,,,,,,,,,,,"bchclinical.systemsupport@nhs.net")
}
