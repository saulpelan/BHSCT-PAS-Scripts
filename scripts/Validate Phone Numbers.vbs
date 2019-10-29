# $language = "VBScript"
# $interface = "1.0"



'	This is a validation script that processes the phone numbers in the Basic Details
'	section of CLINiCOM PMI records to ensure they are valid. Non-numeric characters
'	will be removed, including spaces. 
'
'		Example:	02896 555 555 (MUM)
'		becomes		02896555555
'
'	The script validates both the "Mobile" and "Phone" fields.



Sub Main()
        num1 = crt.Screen.Get(11, 55, 11, 76)
        num2 = crt.Screen.Get(12, 55, 12, 76)
        crt.Screen.Send ("1" & Chr(13))
        GoToField (24)
        crt.Screen.Send (StripNonNumerics(num1) & Chr(13))
        crt.Screen.Send (StripNonNumerics(num2) & Chr(13))
        GoToField (43)
        crt.Screen.Send ("Y" & Chr(13))
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
