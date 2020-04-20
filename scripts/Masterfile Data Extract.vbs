# $language = "VBScript"
# $interface = "1.0"



'	Script name: Masterfile Data Extract
'	Author: Saul Pelan
'
'	┌Description────────────────────────────────────────────────────────────────────┐
'	│ This script will extract all records from a masterfile's superhelp list and	│
'	│ export it to an Excel spreadsheet.						│
'	└───────────────────────────────────────────────────────────────────────────────┘
'
'	┌Instructions───────────────────────────────────────────────────────────────────┐
'	│ Open the superhelp list in a masterfile function, ensure the first page of	│
'	│ records is displaying, and then run the script. Don't press anything until the│
'	│ script has finished scrolling through the masterfile. There is a timeout that	│
'	│ is denoted by "timeoutSeconds" below. If the script hasnt finished scrolling	│
'	│ within that amount of sessions it will terminate. This can be adjusted if	│
'	│ necessary.									│
'       └───────────────────────────────────────────────────────────────────────────────┘



scriptName = "Masterfile Data Extract"
errorState = False
sleepMilliseconds = 1
timeoutSeconds = 180
test = False
Dim data

Sub Main()
	If Not crt.Session.Connected Then
		Exit Sub
	End If
	
	Select Case GetFunctionTitle()

		'Function LCM
		'Had to add a check for "Location code masterfile" as this is the function title on BCH PAS. Not sure how the function seems to have been renamed on BCH, but here we are.
		Case "R e p o r t   T o   L o c a t i o n   M a s t e r   F i l e", "L o c a t i o n   C o d e   M a s t e r   F i l e"

			'Check that superhelp is open
			If GetText(7, 44, 7, 60) = "Code  Description" Then
				ExcelDump "Report To Location", GetEntriesFromScreen(9, 44, 20, 79), 4
			Else
				Msg "Could not locate superhelp list."
				Exit Sub
			End If
		
		'Function ZBF
		Case "U s e r   D e f i n e d   F u n c t i o n   S e t s"

			'Check that superhelp is open
			If GetText(4, 52, 4, 64) = "Function Sets" Then
				ExcelDump "Function Set", GetEntriesFromScreen(5, 36, 22, 79), 3
			Elseif GetText(4, 35, 4, 47) = "Function Code" Then
				ExcelDump "Function", GetEntriesFromScreen(5, 22, 22, 60), 3
			Else
				Msg "Could not locate superhelp list."
				Exit Sub
			End If

		'Function ZAM
		Case "C r e a t e   &   A m e n d   A c c o u n t"

			'Check that superhelp is open
			If GetText(5, 53, 5, 60) = "Username" Then
				ExcelDump "User", GetEntriesFromScreen(6, 34, 22, 79), 12
			Else
				Msg "Could not locate superhelp list."
				Exit Sub
			End If

		'Function CLM
		Case "O P   C l i n i c   M a s t e r   F i l e"

			'Check that superhelp is open
			If GetText(5, 47, 5, 67) = "Code      Description" Then
				ExcelDump "Clinic", GetEntriesFromScreen(7, 47, 22, 79), 8
			Else
				Msg "Could not locate superhelp list."
				Exit Sub
			End If

		'Function DDE
		Case "D o c u m e n t   D e t a i l   M a s t e r   F i l e"

			'Check that superhelp is open
			If GetText(11, 38, 11, 60) = "Document    Description" Then
				ExcelDump "Document", GetEntriesFromScreen(13, 38, 21, 79), 10
			Else
				Msg "Could not locate superhelp list."
				Exit Sub
			End If

		'Function DMM
		Case "D o c t o r   M a s t e r   F i l e"

			'Check that superhelp is open
			If GetText(5, 47, 5, 67) = "Code      Description" Then
				ExcelDump "Doctor", GetEntriesFromScreen(7, 47, 22, 79), 8
			Else
				Msg "Could not locate superhelp list."
				Exit Sub
			End If

		'Function IDM
		Case "D i s c h a r g e   A w a i t e d   R e a s o n   M F"

			'Check that superhelp is open
			If GetText(9, 17, 9, 41) = "Discharge Awaited Reasons" Then
				ExcelDump "Discharge Awaited Reason", GetEntriesFromScreen(11, 17, 18, 79), 6
			Else
				Msg "Could not locate superhelp list."
				Exit Sub
			End If

		Case Else
			Msg "This function or masterfile is not yet supported for superhelp list extraction. Contact Saul Pelan to request it."
	End Select
End Sub

Function GetEntriesFromScreen(ByVal topleftY, ByVal topleftX, ByVal bottomrightY, ByVal bottomrightX)
	Set entries = CreateObject("System.Collections.ArrayList")
	ctime = Timer
	prevtext = ""
	continue = True
	Do While continue
		If Timer - ctime < timeoutSeconds Then
			datalist = crt.Screen.Get2(topleftY, topleftX, bottomrightY, bottomrightX)
			If prevtext = datalist Then
				
			End If
			If prevtext <> datalist Then
				prevtext = datalist
				datalist = Split(datalist, VbCrLf)
				For Each data In datalist
					If data <> "" Then
						entries.Add data
					End If
				Next
			Else
				continue = False
			End If
			SendSpecial("VT_NEXT_SCREEN")
			ctime1 = Timer
			Do While prevtext = crt.Screen.Get2(topleftY, topleftX, bottomrightY, bottomrightX)
				If Timer - ctime1 > 2 Then
					Exit Do
				End If
			Loop
		Else
			Msg "Timed out"
			Exit Do
		End If			
	Loop
	Set GetEntriesFromScreen = entries
End Function

Function GetText(ByVal row1, ByVal col1, ByVal row2, ByVal col2)
	GetText = Trim(crt.Screen.Get(row1, col1, row2, col2))
End Function

Function GetStatusText()
	GetStatusText = GetText(24,1,24,80)
End Function

Function SendSpecial(ByVal text)
	If NOT errorState Then
		crt.Screen.SendSpecial text
		crt.Sleep(sleepMilliseconds)
		If InStr(UCase(GetStatusText), "ERROR") Then
			errorState = True
		End If
	End If
End Function

Function GetFunctionTitle()
	GetFunctionTitle = GetText(1,1,1,80)
End Function

Function GetFunctionSubtitle()
	GetFunctionSubtitle = GetText(2,1,2,41)
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

Function ExcelDump(ByVal desc, ByVal list, ByVal codeLength)
	Set app = CreateObject("Excel.Application")
	Set wb = app.Workbooks.Add()
	Set ws = wb.Worksheets(1)
	Set window = wb.Windows(1)
	ws.Name = "Masterfile Report"

	i = 1
	For Each record In list
		ws.Range("A" & i) = Trim(Left(record, codeLength))
		ws.Range("B" & i) = Trim(Right(record, Len(record) - codeLength))
		i = i + 1
	Next

	ws.Columns("A:B").AutoFit
	
	wb.SaveAs "T:\PM_HealthRec\S_PASSpt\PAS Masterfile Reports\Summaries\" & GetActiveHospital() & "\Clinicom " & GetActiveHospital() & " " & desc & " Masterfile Summary " & Replace(Date, "/", "-")
	app.Visible = True
	Msg "Masterfile report has been saved as: " & VbCrLf & VbCrLf & wb.FullName
End Function

Sub Db(ByVal text)
	If test Then
		MsgBox text
	End If
End Sub

Sub Msg(ByVal text)
	MsgBox "** " & scriptName & " **" & VbCrLf & text
End Sub
