Class IO

	' Sends string <text> to the host. If <sendReturn> is True, sends
	' a carriage return.
	Sub SendText(ByVal text, ByVal sendReturn)
		If sendReturn Then
			text = text + Chr(13)
		End If
		crt.Screen.Send text
	End Sub
	

	' Sends string <text> to the host and waits <timeout> seconds
	' for a cursor update.
	' Returns True if a cursor update is detected within the timeout
	' seconds specified, or False if no cursor update is detected
	' within the specified timeout seconds.

	Function SendTextAndAwaitCursor(ByVal text, ByVal timeout)
		crt.Screen.Send text
		SendTextAndAwaitCursor = crt.Screen.WaitForCursor(timeout) = -1
	End Function


	' Sends string <text> one character at a time, ensuring each character
	' is acknowledged by the host before sending the next.
	' Returns True if the entire string is sent successfully, or False if
	' a character fails to send or if there is an Error message on the
	' screen.

	Function SendCarefully(ByVal text)
		SendCarefully = True
		If Not IsError Then
			For Each char In Split(text, "")
				If Not SendTextAndAwaitCursor(char, 5) Then
					SendCarefully = False
					Exit For
				End If
			Next
		Else
			SendCarefully = False
		End If
	End Function


	' Returns True if the status text on row 24 contains an error
	' message, or False if not.

	Function IsError()
		IsError = InStr(GetStatusText(), "ERROR") = 1
	End Function


	' Returns the status text or End User Help message at row 24.

	Function GetStatusText()
		GetStatusText = Trim(crt.Screen.Get(24, 1, 24, 80))
	End Function
	

	' Sends the character to the host that requests the cursor to be moved up
	' 
	' Returns True if a cursor update is detected or False if the cursor position
	' change is not detected.

	Function MoveCursorUp()
		crt.Screen.SendSpecial("VT_CURSOR_UP")
		MoveCursorUp = crt.Screen.WaitForCursor(1) = -1
	End Function


	' Sends the character to the host that requests the cursor to be moved left
	' 
	' Returns True if a cursor update is detected or False if the cursor position
	' change is not detected.

	Function MoveCursorLeft()
		crt.Screen.SendSpecial("VT_CURSOR_LEFT")
		MoveCursorLeft = crt.Screen.WaitForCursor(1) = -1
	End Function

	
	' Sends the character to the host that requests the cursor to be moved down
	' 
	' Returns True if a cursor update is detected or False if the cursor position
	' change is not detected.

	Function MoveCursorDown()
		crt.Screen.SendSpecial("VT_CURSOR_DOWN")
		MoveCursorRight = crt.Screen.WaitForCursor(1) = -1
	End Function


	' Sends the character to the host that requests the cursor to be moved right
	' 
	' Returns True if a cursor update is detected or False if the cursor position
	' change is not detected.

	Function MoveCursorRight()
		crt.Screen.SendSpecial("VT_CURSOR_RIGHT")
		MoveCursorRight = crt.Screen.WaitForCursor(1) = -1
	End Function

	
	' Sends the cursor as far to the right as it will go.

	Sub CursorEnd()
		Do While MoveCursorRight 
		Loop
	End Sub


	' Sends the cursor as far to the left as it will go.

	Sub CursorHome()
		Do While MoveCursorLeft 
		Loop
	End Sub


	' Attempts to open the superhelp menu.
	'
	' Returns True if the superhelp menu is open or False if the superhelp menu
	' is not detected.

	Function OpenSuperHelp()
		OpenSuperHelp = False
		screen0 = crt.Screen.Get2(1, 1, crt.Screen.Rows, crt.Screen.Columns)
		crt.Screen.SendSpecial("VT_F14")
		If WaitForScreenUpdate(screen0, 1) Then
			If Not IsError Then
				OpenSuperHelp = True
			End If
		End If
	End Function


	' Waits <timeout> seconds to detect a change in the screen from <screenText>.
	' <screenText> should be a representation of the entire screen including CrLf's for
	' this to work, where the entire screen means Row 1, Column 1 to #Rows, #Columns as
	' the screen size may change.
	' 
	' Returns True (immediately) if a screen update is detected within the time given or 
	' False if the screen does not change within the time given.

	Function WaitForScreenUpdate(ByVal screenText, ByVal timeout)
		WaitForScreenUpdate = False
		t0 = Timer
		Do While Timer - t0 < timeout 
			If screenText <> crt.Screen.Get2(1, 1, crt.Screen.Rows, crt.Screen.Columns) Then
				WaitForScreenUpdate = True
				Exit Do
			End If
		Loop
	End Function
	

	' Gets the name/title of the currently open Function.
	' If <removeSpacing> is True then the spacing between characters will be removed.
	' For example: 
	' 		"C r e a t e   &   A m e n d   A c c o u n t"
	' 	becomes	"Create & Amend Account"
	'
	' Returns the title of the currently open Function, with letter spacing optionally
	' removed.

	Function GetTitle(ByVal removeSpacing)
		If GetState() = 5 Then
			title = Trim(crt.Screen.Get(1, 1, 1, crt.Screen.Columns))
			If removeSpacing Then
				GetTitle = Replace(Replace(Replace(title, "   ", "@"), " ", ""), "@", " ")
			Else
				GetTitle = title
			End If
		End If
	End Function


	' Gets the current state of the user's screen.
	
	' Returns	0	-	if terminal/device login screen is active
	'		1	-	if system login screen is active
	'		2	-	if user logged out screen is active
	'		3	-	if Function selection menu is active
	'		4	-	if Function Set selection menu is active
	'		5	-	if user is in a Function

	Function GetState()
		If crt.Screen.Get(2, 1, 2, 5) = "HP-UX" And crt.Screen.Get(4, 1, 4 ,6) = "login:" Then
			GetState = 0
		ElseIf crt.Screen.Get(1, 41, 1, 49) = "/dev/pts/" Then
			If crt.Screen.Get(4, 11, 4, 20) = "Username :" And _
			  crt.Screen.Get(5, 11, 5, 20) = "Password :" Then
				GetState = 1
			ElseIf crt.Screen.Get(4, 7, 4, 20) = "Function Set :" And _
			  crt.Screen.Get(5, 11, 5, 20) = "Function :" And _
			  crt.Screen.Get(8, 31, 8, 49) = "Available Functions" Then
				GetState = 3
			ElseIf crt.Screen.Get(4, 7, 4, 20) = "Function Set :" And _
			  crt.Screen.Get(8, 29, 8, 51) = "Available Function Sets" Then
				GetState = 4
			End If
		ElseIf crt.Screen.Get(19, 33, 19, 48) = "User Logged Out." Then
			GetState = 2
		Else
			GetState = 5
		End If			
	End Function
End Class
