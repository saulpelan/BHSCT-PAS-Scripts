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
		'crt.Screen.Synchronous = True
		crt.Screen.Send text
		SendTextAndAwaitCursor = crt.Screen.WaitForCursor(timeout)
		'crt.Screen.Synchronous = False
	End Function


	' Sends string <text> one character at a time, ensuring each character
	' is acknowledged by the host before sending the next.
	' Returns True if the entire string is sent successfully, or False if
	' a character fails to send or if there is an Error message on the
	' screen.

	Function SendCarefully(ByVal text)
		If InStr(GetStatusText(), "ERROR") = 1 Then
			SendCarefully = True
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


	' Returns the status text or End User Help message at row 24

	Function GetStatusText()
		GetStatusText = Trim(crt.Screen.Get(24, 1, 24, 80))
	End Function
End Class

