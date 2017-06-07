global OutlookPIDName := "Outlook.Application.14"

PasteValues(values:="")	;was my favourite subroutine ;) sac- recommend you put "Sleep, 250" before any GoSub PasteValues if it doesn't work. Think the clipboard can't keep up with lots of SendInput statements
{
    If !(values)
        values:=Clipboard
	StringReplace, OutputMe, values,`r,,A
	SendInput % RegexReplace(OutputMe, "^\s+$")	;trim whitespace including enters (sac- was "^\s+|\s+$")
}

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

ClickReturn(clickX, clickY)	;clicks mouse at coords but returns mouse to previous position after click
{
	BlockInput, Mouse	;blocks keypresses and mouse movements whilst mouse movement commands used
	MouseGetPos, currentX, currentY
	MouseMove, %clickX%, %clickY%, 0
	Click
	MouseMove, %currentX%, %currentY%, 0
}

startExe(path, ProcessName)
{
	run % path
	if ErrorLevel
	{
		Msgbox Couldn't run %path%
		return False
	}
	else
	{
		Process, Wait, %ProcessName%
		If ErrorLevel
			Return False
	}
	return True
}

ImageClick(imagePath, offsetX=0, offsetY=0, trans="", variance=0)	;clicks image (with offset) on active screen
{
	Click % ImageFindInActiveWindow(imagePath, offsetX, offsetY, trans, variance)
}

ImageFindInActiveWindow(imageFile, offsetX=0, offsetY=0, trans="", variance=0)	;returns where image is found (searches whole screen) with optional offsets
{
	WinGetPos, TopLeftX, TopLeftY, Width, Height, A	;get active window stats
	If(trans!="")
		trans := "*Trans" . trans . " "	;yes, dear god the spaces make a difference. http://www.autohotkey.com/docs/commands/ImageSearch.htm
	ImageSearch, OutputVarX, OutputVarY, 0, 0, %Width%, %Height%, *%variance% %trans%%imageFile%		;*TransWhite=make white see through *variance=colour variance (0-255)
	return % OutputVarX + offsetX . "," . OutputVarY + offsetY
}

ImageWait(imageFile, variance=0)		;waits for image in active screen
{
	WaitUntilImageFoundInActiveWindow(imageFile, variance)
}

WaitUntilImageFoundInActiveWindow(imageFile, variance=0)
{
    Loop
    {
	    WinGetPos, TopLeftX, TopLeftY, Width, Height, A
	    ImageSearch, OutputVarX, OutputVarY, 0, 0, %Width%, %Height%, *%variance% %imageFile%		;*variance=colour variance (0-255)
	        If ErrorLevel
				Sleep 500
	        Else
				Break
    }
}

Email(Send, To, Subject, BodyClearText, BodyHTML="", BodyFormat=2, From="", CC="", Bcc="", ReplyRecipients="", FlagText="", ReminderTFalse=False, ReminderDateTime="", Importance=1, ReadReceipt=False, DeferredDeliveryDateTime="", AccountToSendFrom="")		;Sends an email. Returns: True/False.
{
	/*
		;use ComObjActive if Outlook is always Open, but check with:
		If !WinExist("ahk_class rctrl_renwnd32")	;if Outlook is closed, Exit
			Return False
	*/
	Item := ComObjCreate(OutlookPIDName).CreateItem(olMailItem:=0)		;olMailItem:=0=="0". Will work even with outlook closed. .14 is Outlook 2010 on Windows 7
	Item.To := To
	Item.Subject := Subject
 	If BodyClearText !=	;default signature is kept if you don't change the .Body property
 		Item.Body := BodyClearText
	If BodyHTML !=
		Item.HTMLBody := BodyHTML
	Item.BodyFormat := BodyFormat	;1=plain text, 2=html, 3=rich text
	Item.SentOnBehalfOfName := 	From		;sets the "From" field. "" is ok as Outlook just uses default account
	Item.CC := CC		;== Recipient1 := sacComObjectReply.Recipients.Add("a@a.com") then Recipient1.Type := 2 ;To=1 Cc=2 Bcc=3
	Item.BCC := Bcc
	Item.FlagRequest := FlagText 	;sets follow up flag *for recipients*! very cool. "" is ok
	Item.ReplyRecipientNames := ReplyRecipients		;sac, depreciated, should use Item.ReplyRecipients.Add("simon crouch")
	If ReminderDateTime !=
		Item.ReminderTime := ReminderDateTime
	Item.ReminderSet := ReminderTFalse
	Item.ReadReceiptRequested := ReadReceipt
	Item.Importance := Importance	;2=high, 1=med, 0=low
	If DeferredDeliveryDateTime !=
		Item.DeferredDeliveryTime := DeferredDeliveryDateTime
	If AccountToSendFrom !=
		Item.SendUsingAccount := Item.Application.Session.Accounts.Item(AccountToSendFrom)	;sac- this took me ages. You can also use a #
	If Send
		Item.Send
	Else
		Item.Display
	Return True
}
