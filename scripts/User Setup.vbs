# $language = "VBScript"
# $interface = "1.0"

Dim hospitalCode
Dim hospitalDesc
Dim excelApp
Dim workbook
Dim worksheet
Dim debug
Dim fcode
Dim password
Dim accessCode
Dim additionalFSets
Dim outlookApp



Sub Main()
		
	debug = False

	set excelApp = GetObject(,"Excel.Application")
	set workbook = excelApp.Workbooks("User Set Up")
	set worksheet = workbook.ActiveSheet

	hospitalCode = GetHospitalCode()	
	

	Select Case GetSystemCode()
		Case "BCH": fcode = "MUA"
		Case Else: fcode = "ZAM"
	End Select

	GeneratePassword()

	OpenFunction fcode

	crt.Screen.Send(GetUsername() & Chr(13)) 'Username Field

	If GetExpirePassword() Then
		crt.Screen.Send(password & "/EXPIRE" & "")
	Else	
		crt.Screen.Send(password)
	End If
	crt.Screen.Send(Chr(13)) 'Password Field
	crt.Screen.Send(GetAccountDesc() & Chr(13)) 'Owner Name Field
	crt.Screen.Send(GetInitials() & Chr(13)) 'Initials Field
	crt.Screen.Send("90" & Chr(13)) 'Expiry Limit Field
	crt.Screen.Send("YES" & Chr(13)) 'Maintain Password Field
	crt.Screen.Send("NO" & Chr(13)) 'Disable User Field
	If Len(GetUsername()) > 10 Then
		crt.Screen.Send(GetInitials())
	Else
		crt.Screen.Send(GetUsername())
	End If
	crt.Screen.Send(Chr(13)) 'Mailbox Field
	crt.Screen.Send("*" & Chr(13)) 'Printer Group Field
	crt.Screen.Send("*" & Chr(13)) 'Queue Group Field
	crt.Screen.Send(GetPhoneNumber() & Chr(13)) 'Phone Number Field

	crt.Screen.Send("+" & Chr(13)) 'Access Code Field - Enter plus sign to auto generate
	crt.Sleep(50) 'Wait for auto-generated access code to appear on screen
	accessCode = Trim(crt.Screen.Get(13,30,13,32)) 'Grab Auto-generated access code from CLINiCOM
	crt.Screen.Send("PRD" & Chr(13)) 'Directory Field - Set up for production system
	crt.Screen.Send(GetHospitalCode() & Chr(13)) 'Hospital Code Field
	crt.Screen.Send(GetDepartmentCode() & Chr(13)) 'Department Code Field
	crt.Screen.Send("0" & Chr(13)) 'Advanced Mode User Field
	crt.Screen.Send("YES" & Chr(13)) 'Full Menu Display Field
	crt.Screen.Send(GetDefaultFSet() & Chr(13)) 
	crt.Screen.Send(GetDefaultFSetLevel() & Chr(13))
	crt.Screen.Send(Chr(13)) 'Default Queue Field
	crt.Screen.Send("YES" & Chr(13)) 'Enter? Field

	Set addFSets = GetAdditionalFsets()
	For Each fset In addFSets
		crt.Screen.Send(Left(fset,Len(fset)-1) & Chr(13))
		crt.Screen.Send(Right(fset,1) & Chr(13))
		crt.Screen.Send(Chr(13))
		crt.Screen.Send("YES" & Chr(13))
	Next

	crt.Sleep(5000)


	If GetContractRights() = "YES" Then
		OpenFunction "CUI"
		crt.Screen.Send("ADD" & Chr(13))
		crt.Screen.Send(accessCode & Chr(13))
		Select Case GetSystemCode()
			Case "BCH": crt.Screen.Send("BCH")
			Case "MIH": crt.Screen.Send("MIH")
			Case "RGH": crt.Screen.Send("RGHT")
		End Select
		crt.Screen.Send(Chr(13))
		GoToField(6)
		crt.Sleep(5000)
		crt.Screen.Send("YES" & Chr(13))
	End If

	OpenFunction "UAC"
	crt.Screen.Send(accessCode & Chr(13))
	crt.Screen.Send(GetHCNSearch() & Chr(13))
	crt.Screen.Send(GetHCNCompare() & Chr(13))
	GoToField(5)
	crt.Sleep(5000)
	crt.Screen.Send("YES" & Chr(13))

	OpenFunction "BAL"
	crt.Screen.Send("ADD" & Chr(13))
	crt.Screen.Send(accessCode & Chr(13))
	crt.Screen.Send(GetBookingAccess() & Chr(13))
	GoToField(11)
	If GetClinicManagement() = "All" Then
		crt.Screen.Send("ALL" & Chr(13))
	End If
	GoToField(14)
	crt.Sleep(5000)
	crt.Screen.Send("YES" & Chr(13))

	owbAccess = GetOPWLAccess()
	If owbAccess <> "NONE" Then
		OpenFunction "OWB"
		crt.Screen.Send("ADD" & Chr(13))
		crt.Screen.Send(accessCode & Chr(13))
		crt.Screen.Send(owbAccess & Chr(13))
		GoToField(8)
		crt.Sleep(5000)
		crt.Screen.Send("YES" & Chr(13))
	End If
	
	wdaAccess = GetIPWLAccess()
	If owbAccess <> "NONE" Then
		OpenFunction "WDA"
		crt.Screen.Send("ADD" & Chr(13))
		crt.Screen.Send(accessCode & Chr(13))
		crt.Screen.Send(wdaAccess & Chr(13))
		GoToField(6)
		crt.Sleep(5000)
		crt.Screen.Send("YES" & Chr(13))
	End If

	Set outlookApp = GetObject(,"Outlook.Application")
	Set email = outlookApp.CreateItem(0)
	With email
		.To = GetEmailAddress()
		.Subject = GetSystemName()
		.CC = "PAS.Support"
		.HTMLBody = "<body style='padding: 0; margin: 0;'>" & _
			"<div style='position: fixed; top: 0px; left: 0px; background-color: #555555; color: #ccc; margin: 0; padding: 0; height: 100px; width: 100%'>" & _
			"<p style='font-size: 40px; font-family: ""Calibri Light""; text-align: center;'>" & _
			"Patient Administration system" & _
			"</p>" & _
			"<p style='font-size: 18px; font-family: ""Courier New""; text-align: center; color: #ffffff'>" & _
			"Royal Hospitals Trust PAS" & _ 
			"</p>" & _
			"</div>" & _
			"</body>"
		.SentOnBehalfOfName = "PAS.Support"
		.Display
	End With

End Sub

Sub GoToField(ByVal index)
        crt.Screen.SendSpecial ("VT_F9")
        crt.Screen.Send (index & Chr(13))
End Sub

Function OpenFunction(ByVal code)
	If IsMainMenuVisible() Then				'Check if we are already on Main Menu		
		crt.Screen.SendSpecial("VT_F16")
		crt.Screen.Send(code & Chr(13) & Chr(13))		'Type the function code and hit enter twice
	Elseif ExitToMainMenu() Then
		OpenFunction code
	Else
		MsgBox("Error: Could not open function " & code)
	End If
End Function

Function ExitToMainMenu()
	ExitToMainMenu = False
	If IsMainMenuVisible() Then
		ExitToMainMenu = True
	Else
		For i = 1 To 5
			crt.Screen.SendSpecial("VT_F6")
			crt.Sleep(100)
			If IsMainMenuVisible() Then
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
		IsMainMenuVisible = False
	End If
End Function

Function GetTotalRows()
	DoDebug "Getting Total Rows"
	i = 4
	Do While i < 100 AND NOT (IsEmpty(worksheet.Range(fnColumn & i)) AND IsEmpty(worksheet.Range(snColumn & i)) AND IsEmpty(worksheet.Range(dobColumn & i)))
		i = i + 1
	Loop
	DoDebug "Total rows = " & i
	GetTotalRows = i
End Function

Function GetActiveRow()
	DoDebug "Checking active row"
	For i = 4 To GetTotalRows() + 4
		set rng = workbook.ActiveSheet.Range(fnColumn & i & "," & snColumn & i & "," & dobColumn & i)
		If rng.Style = "Normal" Then
			DoDebug "Active row is " & i
			GetActiveRow = i
			Exit For
		End If
	Next
End Function

Sub MarkData(ByVal row, Byval style)
	set rng = workbook.ActiveSheet.Range(fnColumn & row & "," & snColumn & row & "," & dobColumn & row)
	rng.Style = style
End Sub

Sub GeneratePassword()
	Select Case GetSystemCode()
		Case "MIH": password = "MATER19"
		Case "BCH": password = "CITY19"
		Case "RGH": password = "ROYAL19"
	End Select
End Sub

Function GetExpirePassword()
	GetExpirePassword = worksheet.Range("C9") = "YES"
End Function

Function GetForename()
	DoDebug "Looking up forename"
	GetForename = worksheet.Range("C1")
End Function

Function GetSurname()
	DoDebug "Looking up surname"
	GetSurname = worksheet.Range("C2")
End Function

Function GetUsername()
	GetUsername = worksheet.Range("C6")
End Function

Function GetEmployer()
	DoDebug "Looking up employer"
	GetEmployer = worksheet.Range("G1")
End Function

Function GetHospitalCode()
	GetHospitalCode = worksheet.Range("C10")
End Function

Function GetHospitalDesc()
	GetHospitalDesc = worksheet.Range("D10")
End Function

Function GetSystemCode()
	GetSystemCode = worksheet.Range("I9")
End Function

Function GetSystemName()
	GetSystemName = worksheet.Range("H10")
End Function

Sub DoDebug(ByVal text)
	If debug = True Then
		MsgBox text
	End If
End Sub

Function GetAccountDesc()
	GetAccountDesc = worksheet.Range("C7")
End Function

Function GetInitials()
	GetInitials = Left(GetForename(),1) & Left(GetSurname(),1)
End Function

Function GetPhoneNumber()
	GetPhoneNumber = worksheet.Range("C8")
End Function

Function GetDepartmentCode()
	Select Case GetHospitalCode()
		Case "CH": GetDepartmentCode = "CH"
		Case "DH": GetDepartmentCode = ""
		Case "RV": GetDepartmentCode = "MAIN"
		Case "RMH": GetDepartmentCode = "RMH"
		Case "MPH": GetDepartmentCode = "MPH"
		Case "BCH": GetDepartmentCode = "MAIN"
		Case "MIH": GetDepartmentCode = "MIH"
	End Select
End Function

Function GetDefaultFSet()
	GetDefaultFSet = worksheet.Range("C15")
End Function

Function GetDefaultFSetLevel()
	GetDefaultFSetLevel = worksheet.Range("F15")
End Function

Function GetAdditionalFSets()
	DoDebug "Creating ArrayList"
	Set additionalFSets = CreateObject("System.Collections.ArrayList")
	DoDebug "ArrayList Created"
	For i = 19 To 30
		fsetCell = worksheet.Range("D" & i)
		fsetLvlCell = worksheet.Range("E" & i)
		If NOT IsEmpty(fsetCell) Then
			DoDebug "Adding " & fsetCell & " at level " & fsetLvlCell
			additionalFsets.Add(fsetCell & fsetLvlCell)
			DoDebug "Added " & fsetCell & " at level " & fsetLvlCell
		End If
	Next
	DoDebug "Setting value"
	Set GetAdditionalFSets = additionalFSets
	DoDebug "Value set"
End Function

Function GetBookingAccess()
	GetBookingAccess = worksheet.Range("C12")
End Function

Function GetClinicManagement()
	GetClinicManagement = worksheet.Range("F12")
End Function

Function GetOPWLAccess()
	GetOPWLAccess = worksheet.Range("I12")
End Function

Function GetIPWLAccess()
	GetIPWLAccess = worksheet.Range("L12")
End Function

Function GetContractRights()
	GetcontractRights = worksheet.Range("C13")
End Function

Function GetHCNSearch()
	GetHCNSearch = worksheet.Range("F13")
End Function

Function GetHCNCompare()
	GetHCNCompare = worksheet.Range("I13")
End Function

Function GetOverrideEmail()
	GetOverrideEmail = worksheet.Range("F4") = "YES"
End Function

Function GetEmailAddress()
	Select Case GetOverrideEmail()
		Case True: GetEmailAddress = worksheet.Range("F5")
		Case False: GetEmailAddress = worksheet.Range("F3")
	End Select
End Function
