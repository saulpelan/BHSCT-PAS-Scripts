# $language = "VBScript"
# $interface = "1.0"

Sub Main()
	num1 = crt.Screen.Get(11, 55, 11, 76)	'The text representing the value recorded in the "Mobile" field
	num2 = crt.Screen.Get(12, 55, 12, 76)	'The text representing the value recorded in the "Phone" field (for landlines)
	crt.Screen.Send ("1" & Chr(13))	'Send "1" and a return, telling PAS to enter the basic details editing mode
	GoToField (24)	'Go to the "Mobile" field (field #24)
	crt.Screen.Send (StripNonNumerics(num1) & Chr(13))	'Send the stripped version of the number and send a return. The return moves the cursor to the next "Phone" field
	crt.Screen.Send (StripNonNumerics(num2) & Chr(13))	'Now the cursor is on the "Phone" field. Send the stripped version of this number if any and a return to save the value
	GoToField (43)	'Go to the "Enter?" (save changes) field
	crt.Screen.Send ("Y" & Chr(13))	'Type Y for YES and send return
End Sub

Sub GoToField(ByVal index)
	crt.Screen.SendSpecial ("VT_F9")
	crt.Screen.Send (index & Chr(13))
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
