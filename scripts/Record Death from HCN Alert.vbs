# $language = "VBScript"
# $interface = "1.0"



'	V1.1.1
'	Last updated 30/10/19
'	Saul Pelan
'
'	┌Description────────────────────────────────────────────────────────────────────┐
'	│ This script will record patient deaths using the Health & Care number provided│
'	│ in an email alert from the Health & Care index.                               │
'	└───────────────────────────────────────────────────────────────────────────────┘
' 
'	┌Instructions───────────────────────────────────────────────────────────────────┐
'	│ 1.	In MS Outlook, select the HCN alert email so that it is showing in the  │
'	│ 	preview pane.                                                           │
'	│ 2.	Ensure you are connected and logged in to the correct hospital's PAS.   │
'	│ 3.	Run this script in CRT. It will lookup the patient using the H&C number │
'	│	supplied in the email and get the date of death from the H&C compare    │
'       │	feature on CLINiCOM. Then it will open the Record Patient Death         │
'	│	function and fill out the form and leave you to enter "Y" at the end.	│
'	│ 4.	You have 60 seconds to press "Y". Once pressed the script will mark the │
'	│	email as read and move it to the HCN Updates folder. If you take longer │
'	│	than 60 seconds you will have to enter "YES" and move the email manually│
'	│                                                                               │
'	│ Tip: it may be helpful to have both Outlook and CRT visible so you can see    │
'	│ what is happening. Press Win+Left Arrow or Win+Right Arrow on each window to  │
'	│ split them on the screen.                                                     │
'       └───────────────────────────────────────────────────────────────────────────────┘



Dim DoD
Dim emailHosp
Dim activeHosp

Sub Main()
	If GetEmailSubject() = "ALERT - MANUAL UPDATE REQUESTED ON PAS" Then
		If InStr(GetEmailBody(), "Date of Death: ") Then
			hcn = GetHCNFromEmail()
			If GetActiveHospital() <> emailHosp Then
				MsgBox "Active hospital's PAS does not match email. Script stopped."
				Exit Sub
			End If
			If Not (crt.Screen.Get(1,21,1,59) = "R e c o r d   P a t i e n t   D e a t h" And crt.Screen.Get(2,6,2,30) = "Patient Selection Details") Then
				OpenFunction "RPD", "R e c o r d   P a t i e n t   D e a t h"
			End If
			GoToField(6)
			crt.Screen.SendSpecial("VT_REMOVE")
			GoToField(6)				'Go to H+C field
			crt.Screen.Send(hcn & Chr(13))
			crt.Sleep 25
			If crt.Screen.Get(24,1,24,5) = "ERROR" Then
				MsgBox "Error"			'Check for error
				Exit Sub
			End If
			crt.Screen.Send("YES" & Chr(13))
			If crt.Screen.WaitForString("** Died **", 1) Then
				crt.Sleep 25
				DoD = crt.Screen.Get(19,69,19,78)
				If Mid(DoD,3,1) = "/" Then
					crt.Screen.SendSpecial("VT_F6")
					crt.Screen.SendSpecial("VT_F6")
					OpenFunction "RPD", "R e c o r d   P a t i e n t   D e a t h"
					GoToField(6)
					crt.Screen.Send(hcn & Chr(13))
					crt.Screen.Send("NO" & Chr(13) & "NO" & Chr(13) & "YES" & Chr(13) & DoD & Chr(13))
					crt.Screen.Send(GetEmailSender())
					crt.Screen.Send(Chr(13) & Chr(13))	
					crt.Screen.WaitForKey(60)
					crt.Sleep(1)
					If UCase(crt.Screen.Get(22,12,22,12)) = "Y" Then
						crt.Screen.Send(Chr(13))
						MoveEmail()
					Else
						Exit Sub
					End If						
				End If
			Else
			End If
		Else
			MsgBox "This is not a revision of date of death. Please check email."
			Exit Sub
		End If
	Else
		MsgBox "This is not a manual update request from HCN. Please check email." & vbCrLF & "Subject: " & GetEmailSubject() & vbCrLf & "From: " & GetEmailSender()
	End If
End Sub

Sub GoToField(ByVal index)
        crt.Screen.SendSpecial ("VT_F9")
        crt.Screen.Send (index & Chr(13))
End Sub

Function OpenFunction(ByVal code, ByVal title)
	OpenFunction = False
	If IsMainMenuVisible() Then				'Check if we are already on Main Menu
		crt.Screen.SendSpecial("VT_F16")			'(F10 button)
		crt.Screen.Send(code & Chr(13) & Chr(13))		'Type the function code and hit enter twice
	Elseif crt.Screen.WaitForString("/dev/pts/", 1) Then	'If not, check if Main Menu is opening
		crt.Screen.SendSpecial("VT_F16") '(F10 button)		'(F10 button)
		crt.Screen.Send(code & Chr(13) & Chr(13))		'Type the function code and hit enter twice
	Else							'If not then 
		If crt.Screen.WaitForString(title,1) Then		'Check if the desired function is already in the process of opening
			OpenFunction = True					'If it is then return true
		ElseIf ExitToMainMenu() Then				'If not, attempt to exit to main menu and check we have done so
			OpenFunction code, title				'And Try again to open the function
		Else							'If we werent able to then warn user and terminate this function
			MsgBox "Couldn't navigate to main menu to open function " & code	
		End If
	End If
End Function

Function ExitToMainMenu()
	ExitToMainMenu = False
	If IsMainMenuVisible() Then
		ExitToMainMenu = True
	Else
		For i = 1 To 5
			crt.Screen.SendSpecial("VT_F6")
			If crt.Screen.WaitForString("/dev/pts/",1) Then
				ExitToMainMenu = True
				Exit For
			End If
		Next
	End If
End Function

Function IsMainMenuVisible()
	If crt.Screen.Get(1,41,1,49) = "/dev/pts/" Then
		IsMainMenuVisible = True
	Else
		IsMainMenuVisible = crt.Screen.WaitForString("/dev/pts/", 1)
	End If
End Function

Function GetHCNFromEmail()
	set obj = GetObject(,"Outlook.Application")
	set email = obj.ActiveExplorer.Selection(1).GetInspector.WordEditor
	set words = email.Words
	Dim hcn
	If words.Item(39) = "MATER " Then
		hcn = words.Item(46) & words.Item(47) & words.Item(48)
		emailHosp = "MIH"
	Elseif words.Item(39) = "BELFAST " Then 
		hcn = words.Item(47) & words.Item(48) & words.Item(49)
		emailHosp = "BCH"
	Elseif words.Item(39) = "ROYAL " Then
		hcn = words.Item(47) & words.Item(48) & words.Item(49)
		emailHosp = "RGH"
	Else
		hcn = "ERROR"
	End If
	GetHCNFromEmail = StripNonNumerics(hcn)
End Function

Function MoveEmail()
	set obj = GetObject(,"Outlook.Application")
	set ns = obj.GetNameSpace("MAPI")
	set email = obj.ActiveExplorer.Selection(1)
	email.UnRead = False
	email.Move ns.Folders("PAS.Support-SM").Folders("Inbox").Folders("HCN Updates")
End Function

Function GetEmailSender()
	GetEmailSender = GetEmailSenderName() & " <" & GetEmailSenderAddress() & ">" 
End Function

Function GetEmailSenderName()
	set obj = GetObject(,"Outlook.Application")
	set email = obj.ActiveExplorer.Selection(1)
	GetEmailSenderName = email.SenderName
End Function

Function GetEmailSubject()
	set obj = GetObject(,"Outlook.Application")
	set email = obj.ActiveExplorer.Selection(1)
	GetEmailSubject = email.Subject
End Function

Function GetEmailBody()
	set obj = GetObject(,"Outlook.Application")
	set email = obj.ActiveExplorer.Selection(1)
	GetEmailBody = email.Body
End Function

Function GetEmailSenderAddress()
	set obj = GetObject(,"Outlook.Application")
	set email = obj.ActiveExplorer.Selection(1)
	If email.SenderEmailType = "SMTP" Then
		GetEmailSenderAddress = email.SenderEmailAddress
	Else
		GetEmailSenderAddress = email.Sender.GetExchangeUser().PrimarySmtpAddress
	End If
End Function

Function StripNonNumerics(ByVal num)
    newString = ""
    For i = 1 To Len(num)
        If IsNumeric(Mid(num, i, 1)) Then
        newString = newString & Mid(num, i, 1)
        End If
    Next
    If newString = "" Then
        newString = " "
    End If
    StripNonNumerics = newString
End Function

Function GetActiveHospital()
	If InStr(crt.Window.Caption, "Belfast City") Then
		GetActiveHospital = "BCH"
	Elseif InStr(crt.Window.Caption, "Mater") Then
		GetActiveHospital = "MIH"
	Elseif InStr(crt.Window.Caption, "Royal Group") Then
		GetActiveHospital = "RGH"
	Else
		GetActiveHospital = "ERROR"
	End If
End Function
