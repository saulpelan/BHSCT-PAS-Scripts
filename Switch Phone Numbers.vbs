# $language = "VBScript"
# $interface = "1.0"

Sub Main()
  num1 = crt.Screen.Get(11, 55, 11, 76)   'get the text in the mobile fiel
  num2 = crt.Screen.Get(12, 55, 12, 76)   'get the text in the landline fiel
  crt.Screen.Send ("1" & Chr(13))	'enter Edit details mode 
  GoToField (24)	'go to field 24 (mobile field)
  crt.Screen.Send (num2 & Chr(13))	'send the original landline text into the mobile field and a carriage return, moving the cursor to the landline field
  crt.Screen.Send (num1 & Chr(13))	'now the cursor is in the landline field, send the original text from the mobile field and carriage return to save the field
  GoToField (43)	'go to the "Enter?" field
  crt.Screen.Send ("Y" & Chr(13))	'send Y for YES and a carriage return. The switched numbers are now saved
End Sub

Sub GoToField(ByVal index)
  crt.Screen.SendSpecial ("VT_F9")	'CRT sends the predefined VT_F9 function (equates to ANSI escape sequence <escape character>[20~) which on PAS opens the Go To Field option
  crt.Screen.Send (index & Chr(13))	'in the Go To Field menu, send the field number <index> and a carriage return to jump to the field
End Sub
