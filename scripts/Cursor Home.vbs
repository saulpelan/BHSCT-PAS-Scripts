# $language = "VBScript"
# $interface = "1.0"



'	Script name: Cursor Home
'	Author: Saul Pelan
'
'	┌Description────────────────────────────────────────────────────────────────────┐
'	│ This script attempts to send the terminal cursor to the start of the current	│
'	│ field, simulating the "Home" keyboard key.					│
'	└───────────────────────────────────────────────────────────────────────────────┘
'
'	┌Instructions───────────────────────────────────────────────────────────────────┐
'	│ In CRT, map a key (Home key recommended) to this script. Press the mapped key	│
'	│ when editing a field on CLINiCOM.						│
'       └───────────────────────────────────────────────────────────────────────────────┘

scriptName = "Cursor Home"
col = crt.Screen.CurrentColumn

' User Defined Options 
A0 = 500	' Time (in milliseconds) to wait after sending the cursor manipulation string before 
		' 	terminating the script. NB: 1000 milliseconds = 1 second
		'	Values: any integer (whole number)


Sub Main()

	' This instructs the terminal emulator to send the character that modifies the cursor position.
	' The way CLINiCOM handles this character is as follows:
	' 
	' 	If cursor position is at the very start of a field: move the cursor to the end.
	' 	If cursor is at any other position in a field: move it the cursor to the start.
	'
	'	Some fields such as core system/menu fields and some entire functions such as AMS functions
	'	won't accept this cursor position modifier. Most non-AMS functions will permit it, with the
	'	exception of some special fields/functions such as the paragraph editor in Paragraph 
	'	Detail Masterfile.
	crt.Screen.SendSpecial("VT_SELECT")

	' Wait for the cursor position to update, timing out after the number of milliseconds defined in option A0
	If WaitForCursorTimeout(A0) Then
		newcol = crt.Screen.CurrentColumn

		' Check to see if the cursor is now to the right of its original position, and if it is, send the 
		' cursor position modifying character again in order to send it to the start of the field.
		If newcol > col Then
			crt.Screen.SendSpecial("VT_SELECT")
		End If
	End If
End Sub

' This function delays the script until either a cursor position update is detected or the function times out
' after the amount of milliseconds specified in option A0
Function WaitForCursorTimeout(ByVal timeoutMillis)
	t1 = Timer
	WaitForCursorTimeout = False
	Do Until Timer - t1 >= timeoutMillis / 1000
		If col <> crt.Screen.CurrentColumn Then
			WaitForCursorTimeout = True
			Exit Do
		End If
	Loop
End Function

Sub SendSpecial(ByVal text)
	If errorState Then 
		Exit Sub
	End If
	crt.Screen.SendSpecial text
End Sub	

' If option A1 is switched on a MsgBox will display text to the user
Sub Msg(ByVal text)
	If A1 Then
		MsgBox "** " & scriptName & " **" & VbCrLf & text
	End If
End Sub
