# $language = "VBScript"
# $interface = "1.0"



'This script switches the phone numbers in the Basic Details section in a CLINiCOM
'PMI record. Sometimes mobile and landline numbers are recorded in the wrong fields.
'
'Example:
'		Mobile No :02890777777
'		Phone     :07444444444 
'	
'	becomes
'		Mobile No :07444444444
'		Phone     :02890777777



Sub Main()
        num1 = crt.Screen.Get(11, 55, 11, 76)
        num2 = crt.Screen.Get(12, 55, 12, 76)
        crt.Screen.Send ("1" & Chr(13))
        GoToField (24)
        crt.Screen.Send (num2 & Chr(13))
        crt.Screen.Send (num1 & Chr(13))
        GoToField (43)
        crt.Screen.Send ("Y" & Chr(13))
End Sub

Sub GoToField(ByVal index)
        crt.Screen.SendSpecial ("VT_F9")
        crt.Screen.Send (index & Chr(13))
End Sub
