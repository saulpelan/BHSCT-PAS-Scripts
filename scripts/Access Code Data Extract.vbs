# $language = "VBScript"
# $interface = "1.0"



'	Script name: Access Code Data Extract
'	Author: Saul Pelan
'
'	┌Description────────────────────────────────────────────────────────────────────┐
'	│ This script interrogates the CLINiCOM PAS for information recorded against a  │
'	│ user access code and generates an Excel spreadsheet with a full description.  │
'	└───────────────────────────────────────────────────────────────────────────────┘
' 
'	┌Intended use───────────────────────────────────────────────────────────────────┐
'	│ CLINiCOM PAS is a VT220 terminal application meaning it is limited to a       │
'	│ textual user interface of 24 rows and 80 columns. Naturally this means that   │
'	│ not all data can be displayed on one page easily and when manually            │
'	│ interrogating a record there are several menus and pages that have to be      │
'	│ navigated for a full in-depth understanding of the record.                    │                                                           │
'	│                                                                               │
'	│ This script acts as an interface between the VT220 display and a modern       │
'	│ spreadsheet. It navigates all the menus required to obtain a full description │
'	│ of a record and extracts data as it goes along and uses the data to generate a│
'	│ Excel spreadsheet.                                                            │
'	│										│
'	│ This script is useful when you are trying to set up access for someone,	│
'	│ while matching the access of an existing user.				│
'	└───────────────────────────────────────────────────────────────────────────────┘
'
'	┌Instructions───────────────────────────────────────────────────────────────────┐
'	│ 1.	When in a function or the main function set menu, simply run this script│
'	│ 	and you will be prompted to enter an access code to report on.		│
'	│ 2.	Enter the access code when prompted, wait a few seconds, and the script	│
'	│ 	will generate a report in Excel spreadsheet format. 			│
'       └───────────────────────────────────────────────────────────────────────────────┘



Dim accessCode
Dim username
Dim userDesc
Dim initials
Dim expiryLimit
Dim expiryDate
Dim maintainPassword
Dim disabled
Dim phone
Dim directory
Dim hospitalCode
Dim department
Dim defaultFunctionSet
Dim defaultFunctionSetLevel
Dim functionSets
Dim pcAccess
Dim balOption
Dim balClinicGroups
Dim balClinics
Dim balManagementOptions
Dim wdaOption
Dim wdaWLGroups
Dim wdaWLs
Dim owbOption
Dim owbWLGroups
Dim owbWLs
Dim contractCode
Dim contractDesc
Dim hcSearch
Dim hcCompare

errorState = False
sleepMilliseconds = 25
test = False


Sub Main()
	If Not crt.Session.Connected Then
		Exit Sub
	End If
	accessCode = UCase(crt.Dialog.Prompt("Enter access code:", "Access Code Report"))
        OpenFunction("ZAU")
	Send(accessCode)
	username = GetText(10, 27, 10, 38)
	If errorState Then
		Exit Sub
	End If
	OpenFunction("ZAM")
	Send(username)
	userDesc = GetText(5, 21, 5, 50)
	initials = GetText(6, 21, 6, 23)
	expiryDate = GetText(6, 65, 6, 72)
	expiryLimit = GetText(7, 21, 7, 22)
	maintainPassword = GetText(7, 65, 7, 67) = "YES"
	disabled = GetText(8, 21, 8, 23) = "YES"
	phone = GetText(10, 21, 10, 36)

	For i = 1 To 10
		Send("")
	Next

	Send(accessCode)
	directory = GetText(14, 30, 14, 45)
	hospitalCode = GetText(15, 30, 15, 33)
	department = GetText(15, 65, 15, 72)
	
	For i = 1 To 8
		Send("")
	Next
	
	Send("Y")
	
	defaultFunctionSet = GetText(5, 18, 5, 20)
	defaultFunctionSetLevel = GetText(5, 73, 5, 73)
	Set functionSets = GetFunctionSets()

	OpenFunction("CWU")
	Send(accessCode)
	pcAccess = GetText(8, 30, 8, 39)
	
	OpenFunction("BAL")
	Send("REVISE")
	If AttemptSend(accessCode) Then
		balOption = GetText(6, 24, 6, 30)
		If balOption <> "BOOK" Then
			Send("")
			Set balClinicGroups = GetBALClinicGroups()
			GoToFieldOcc 8
			Set balClinics = GetBALClinics()
		End If
		GoToPage 2
		balManagementOptions = NOT GetText(5, 4, 5, 77) = ""
	Else
		balOption = "N/A"
	End If

	OpenFunction("WDA")
	Send("LIST")
	If AttemptSend(accessCode) Then
		wdaOption = GetText(6, 19, 6, 25)
		If wdaOption <> "BOOK" Then
			Set wdaWLGroups = GetWDAWLGroups()
			Set wdaWLs = GetWDAWLs()
		End If
	Else
		wdaOption = "N/A"
	End If

	OpenFunction("OWB")
	Send("LIST")
	If AttemptSend(accessCode) Then
		owbOption = GetText(8, 24, 8, 30)
		If owbOption <> "FULL" Then
			Set owbWLGroups = GetOWBWLGroups()
			Set owbWLs = GetOWBWLs()
		End If
	Else
		owbOption = "N/A"
	End If

	OpenFunction("CUI")
	Send("LIST")
	If AttemptSend(accessCode) Then
		contractCode = GetText(8, 29, 8, 32)
		contractDesc = GetText(8, 36, 8, 76)
	Else
		contractCode = "N/A"
	End If
	
	OpenFunction("UAC")
	hcSearch = False
	hcCompare = False
	If AttemptSend(accessCode) Then
		hcSearch = GetText(7, 30, 7, 32) = "YES"
		hcCompare = GetText(9, 30, 9, 32) = "YES"
	End If

	ExitToMenu()
	
	GenerateReport()

End Sub

Sub GoToField(ByVal index)
        SendSpecial("VT_F9")
        Send(index)
End Sub

Sub GoToFieldOcc(ByVal index)
        SendSpecial("VT_F9")
        Send(index)
	Send("1")	
End Sub

Sub GoToPage(ByVal index)
	SendSpecial("VT_F10")
	Send(index)
End Sub
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

Function IsMainMenuVisible()
	IsMainMenuVisible = crt.Screen.Get(1,41,1,49) = "/dev/pts/"
End Function

Function ExitToMenu()
	ExitToMenu = False
	For i = 1 To 5
		If IsMainMenuVisible() Then
			ExitToMenu = True
			Exit For
		Else
			SendSpecial("VT_F6")
		End If
	Next
End Function

Sub OpenFunction(ByVal code)
	If NOT IsMainMenuVisible() Then
		If NOT ExitToMenu() Then
			errorState = True
			MsgBox "Could not open function: " & code	
			Exit Sub
		End If
	End If
	SendSpecial("VT_F16")
	Send(code)
	Send("")
End Sub

Function SendSpecial(ByVal text)
	If NOT errorState Then
		crt.Screen.SendSpecial text
		crt.Sleep(sleepMilliseconds)
		If InStr(UCase(GetStatusText), "ERROR") Then
			errorState = True
		End If
	End If
End Function

Function Send(ByVal text)
	If NOT errorState Then
		crt.Screen.Send text & Chr(13)
		crt.Sleep(sleepMilliseconds)
		If InStr(UCase(GetStatusText()), "ERROR") Then
			errorState = True
		End If
	End If
End Function

Function AttemptSend(ByVal text)
	If NOT errorState Then
		crt.Screen.Send text & Chr(13)
		crt.Sleep(sleepMilliseconds)
		If InStr(UCase(GetStatusText()), "ERROR") Then
			AttemptSend = False
		Else
			AttemptSend = True
		End If
	End If
End Function

Function GetText(ByVal row1, ByVal col1, ByVal row2, ByVal col2)
	GetText = Trim(crt.Screen.Get(row1, col1, row2, col2))
End Function

Function GetStatusText()
	GetStatusText = GetText(24,1,24,80)
End Function

Function GetFunctionSets()
	lines = Split(Trim(crt.Screen.Get2(15, 3, 21, 77)), VbCrLf)
	Set fsets = CreateObject("System.Collections.ArrayList")
	For Each line In lines
		If Trim(line) <> VbLfCr Then
			words = Split(Trim(line), "  ")
			For Each word In words
				If Trim(word) <> "" Then
					fsets.Add(Trim(word))
				End If
			Next
		End If
	Next
	Set GetFunctionSets = fsets
End Function

Function GetBALClinicGroups()
	Set cgs = CreateObject("System.Collections.ArrayList")
	clinicGroups = Split(crt.Screen.Get2(11, 7, 15, 14), VbCrLf)
	For Each clinicGroup In clinicGroups
		'If Trim(clinicGroup) <> "" Then
			cgs.Add(Trim(clinicGroup))
		'End If
	Next
	Set GetBALClinicGroups = cgs
End Function

Function GetBALClinics()
	Set cs = CreateObject("System.Collections.ArrayList")
	clinics = Split(crt.Screen.Get2(17, 7, 21, 14), VbCrLf)
	For Each clinic In clinics
		'If Trim(clinic) <> "" Then
			cs.Add(Trim(clinic))
		'End If
	Next
	Set GetBALClinics = cs
End Function

Function GetWDAWLGroups()
	Set wlgs = CreateObject("System.Collections.ArrayList")
	wlGroups = Split(crt.Screen.Get2(8, 19, 12, 26), VbCrLf)
	For Each wlGroup In wlGroups
		'If Trim(wlGroup) <> "" Then
			wlgs.Add(Trim(wlGroup))
		'End If
	Next
	Set GetWDAWLGroups = wlgs
End Function

Function GetWDAWLs()
	Set wls = CreateObject("System.Collections.ArrayList")
	wlists = Split(crt.Screen.Get2(14, 19, 18, 24), VbCrLf)
	For Each wl In wlists
		'If Trim(wl) <> "" Then
			wls.Add(Trim(wl))
		'End If
	Next
	Set GetWDAWLs = wls
End Function

Function GetOWBWLGroups()
	Set wlgs = CreateObject("System.Collections.ArrayList")
	wlGroups = Split(crt.Screen.Get2(17, 5, 20, 12), VbCrLf)
	For Each wlGroup In wlGroups
		'If Trim(wlGroup) <> "" Then
			wlgs.Add(Trim(wlGroup))
		'End If
	Next
	Set GetOWBWLGroups = wlgs
End Function

Function GetOWBWLs()
	Set wls = CreateObject("System.Collections.ArrayList")
	wlists = Split(crt.Screen.Get2(12, 5, 15, 12), VbCrLf)
	For Each wl In wlists
		'If Trim(wl) <> "" Then
			wls.Add(Trim(wl))
		'End If
	Next
	Set GetOWBWLs = wls
End Function

Function GenerateReport()
	Set app = CreateObject("Excel.Application")
	Set wb = app.Workbooks.Add()
	Set ws = wb.Worksheets(1)
	Set window = wb.Windows(1)
	window.Caption = accessCode & " - " & username & " - " & userDesc
	window.DisplayGridlines = False

	ws.Name = accessCode
	
	colourValue = 125

	With ws.Range("C1:C100,D1:D100,G1:G100,K1:K100")
		.Font.Color = RGB(colourValue,colourValue,colourValue)
		.HorizontalAlignment = -4131
	End With


	
	ws.Range("A1") = "User Details"
	ws.Range("A1").Font.Bold = True
	ws.Range("A1:G1").Merge
	
	ws.Range("A2") = "Username"
	ws.Range("B2") = ":"
	ws.Range("C2") = username
	
	If disabled Then
		ws.Range("D2") = "** Disabled **"
		ws.Range("A1,D2").Font.Color = vbRed
	End If

	ws.Range("E2") = "Owner"
	ws.Range("F2") = ":"
	ws.Range("G2") = userDesc

	ws.Range("A3") = "Access Code"
	ws.Range("B3") = ":"
	ws.Range("C3") = accessCode

	ws.Range("E3") = "Initials"
	ws.Range("F3") = ":"
	ws.Range("G3") = initials
	
	ws.Range("A4") = "Phone Number"
	ws.Range("B4") = ":"
	ws.Range("C4") = phone


	
	ws.Range("A6") = "Password Details"
	ws.Range("A6").Font.Bold = True
	ws.Range("A6:G6").Merge
	
	ws.Range("A7") = "Password set on"
	ws.Range("B7") = ":"
	ws.Range("C7") = CDate(CDate(expiryDate)-expiryLimit)

	ws.Range("E7") = "Expires on"
	ws.Range("F7") = ":"
	ws.Range("G7") = CDate(expiryDate)
	If CDate(Date) >= CDate(expiryDate) Then
		ws.Range("A6,G6").Font.Color = vbRed
		ws.Range("E7") = "Expired on"
	End If

	ws.Range("A8") = "Maintain Pwd?"
	ws.Range("B8") = ":"
	ws.Range("C8") = maintainPassword

	ws.Range("E8") = "Expiry Limit"
	ws.Range("F8") = ":"
	ws.Range("G8") = expiryLimit



	ws.Range("A10") = "Access Details"	
	ws.Range("A10").Font.Bold = True
	ws.Range("A10:G10").Merge

	ws.Range("A11") = "Directory"
	ws.Range("B11") = ":"
	ws.Range("C11") = directory

	ws.Range("E11") = "Hospital Code"
	ws.Range("F11") = ":"
	ws.Range("G11") = hospitalCode
	
	ws.Range("A12") = "PatientCentre"
	ws.Range("B12") = ":"
	ws.Range("C12") = pcAccess

	ws.Range("E12") = "Department Code"
	ws.Range("F12") = ":"
	ws.Range("G12") = department

	ws.Range("A13") = "H+C Search"
	ws.Range("B13") = ":"
	ws.Range("C13") = hcSearch
	
	ws.Range("E13") = "H+C Compare"
	ws.Range("F13") = ":"
	ws.Range("G13") = hcCompare

	ws.Range("A14") = "Contract ID"
	ws.Range("B14") = ":"
	ws.Range("C14") = contractCode	
	ws.Range("D14") = contractDesc
	ws.Range("D14:G14").Merge


	
	ws.Range("A16") = "BAL OP Booking Access"
	ws.Range("A16").Font.Bold = True
	ws.Range("A16:G16").Merge

	ws.Range("A17") = "All Clinics"
	ws.Range("B17") = ":"
	ws.Range("C17") = balOption

	ws.Range("E17") = "Clin/Rec Mgmt"
	ws.Range("F17") = ":"
	ws.Range("G17") = balManagementOptions

	If Not IsEmpty(balClinicGroups) Then
		ws.Range("A18") = "Clinic Groups"
		ws.Range("B18") = ":"
		ws.Range("C18") = balClinicGroups.Item(0)
		ws.Range("C19") = balClinicGroups.Item(1)
		ws.Range("C20") = balClinicGroups.Item(2)
		ws.Range("C21") = balClinicGroups.Item(3)
		ws.Range("C22") = balClinicGroups.Item(4)
	End If

	If Not IsEmpty(balClinics) Then
		ws.Range("E18") = "Clinic Codes"
		ws.Range("F18") = ":"
		ws.Range("G18") = balClinics.Item(0)
		ws.Range("G19") = balClinics.Item(1)
		ws.Range("G20") = balClinics.Item(2)
		ws.Range("G21") = balClinics.Item(3)
		ws.Range("G22") = balClinics.Item(4)
	End If




	ws.Range("A24") = "OWB OP W/L Booking Access"
	ws.Range("A24").Font.Bold = True
	ws.Range("A24:G24").Merge
	
	ws.Range("A25") = "All OP W/Ls"
	ws.Range("B25") = ":"
	ws.Range("C25") = owbOption

	If Not IsEmpty(owbWLGroups) then
		ws.Range("A26") = "OP W/L Groups"
		ws.Range("B26") = ":"
		ws.Range("C26") = owbWLGroups.Item(0)
		ws.Range("C27") = owbWLGroups.Item(1)
		ws.Range("C28") = owbWLGroups.Item(2)
		ws.Range("C29") = owbWLGroups.Item(3)
		ws.Range("C30") = owbWLGroups.Item(4)
	End If

	If Not IsEmpty(owbWLs) Then
		ws.Range("E26") = "OP W/L Codes"
		ws.Range("F26") = ":"
		ws.Range("G26") = owbWLs.Item(0)
		ws.Range("G27") = owbWLs.Item(1)
		ws.Range("G28") = owbWLs.Item(2)
		ws.Range("G29") = owbWLs.Item(3)
		ws.Range("G30") = owbWLs.Item(4)
	End If



	ws.Range("A32") = "WDA W/L Data Access"
	ws.Range("A32").Font.Bold = True
	ws.Range("A32:G32").Merge
	
	ws.Range("A33") = "All W/Ls"
	ws.Range("B33") = ":"
	ws.Range("C33") = wdaOption

	If Not IsEmpty(wdaWLGroups) then
		ws.Range("A34") = "W/L Groups"
		ws.Range("B34") = ":"
		ws.Range("C34") = wdaWLGroups.Item(0)
		ws.Range("C35") = wdaWLGroups.Item(1)
		ws.Range("C36") = wdaWLGroups.Item(2)
		ws.Range("C37") = wdaWLGroups.Item(3)
		ws.Range("C38") = wdaWLGroups.Item(4)
	End If
	
	If Not IsEmpty(wdaWLs) then
		ws.Range("E34") = "W/L Codes"
		ws.Range("F34") = ":"
		ws.Range("G34") = wdaWLs.Item(0)
		ws.Range("G35") = wdaWLs.Item(1)
		ws.Range("G36") = wdaWLs.Item(2)
		ws.Range("G37") = wdaWLs.Item(3)
		ws.Range("G38") = wdaWLs.Item(4)
	End If


	
	ws.Range("A40") = "Function Set Access"
	ws.Range("A40").Font.Bold = True
	ws.Range("A40:G40").Merge

	ws.Range("A41") = "Default"
	ws.Range("B41") = ":"
	ws.Range("C41") = defaultFunctionSet & "\" & defaultFunctionSetLevel

	ws.Range("A42") = "All"
	ws.Range("B42") = ":"
	
	Set fsRange = ws.Range("C42:D50")
	i = 1

	For Each functionSet in functionSets
		fsRange(i) = functionSet
		i = i + 1
	Next

	ws.Columns("A:K").AutoFit

	ws.PageSetup.LeftHeader = GetActiveHospital() & " PAS Report"
	ws.PageSetup.CenterHeader = username & " - " & accessCode
	ws.PageSetup.RightHeader = "&D &T"
	ws.PageSetup.RightFooter = "Page &P of &N"

	app.Visible = True

End Function

Sub Db(ByVal text)
	If test Then
		MsgBox text
	End If
End Sub
