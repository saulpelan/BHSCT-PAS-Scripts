# $language = "VBScript"
# $interface = "1.0"

Dim hospital
Dim excelApp
Dim workbook
Dim tab
Dim ForenameColumn
Dim SurnameColumn
Dim DoBColumn

Sub Main()
	
	set excelApp = GetObject(, "Excel.Application");
		
	
		set workbook = excellApp.Workbooks(1)
			
		
		
set excelApp = GetObject(,"Excel.Application")
set workbook = excelApp.Workbooks(1)

If InStr(workbook.Name, "Mater") > 0 Then
hospital = "MIH"
Elseif InStr(workbook.Name, "RGH") > 0 Then
hospital = "RGH"
Elseif InStr(workbook.Name, "BCH") > 0 Then
hospital = "BCH"
Else
MsgBox "Could not determine hospital from workbook name [" & workbook.Name & "]"
Exit Sub
End If	

For i = 1 To crt.GetTabCount
If InStr(crt.GetTab(i).Caption, hospital) Then
	set tab = crt.GetTab(i)
	Exit For
End If
Next

If IsEmpty(tab) OR tab.Screen.Get(1,41,1,49) <> "/dev/pts/" Then
MsgBox "Please log in to the relevant production system with your username and password and restart the script at the main function menu."	
Exit Sub
End If

End If 		

MsgBox "Working on Data Quality Report: " & workbook.Name & vbCrLf & "Working on PAS session: " & tab.Caption
MsgBox "Hi"



End Sub

Sub GoToField(ByVal index)
  crt.Screen.SendSpecial ("VT_F9")
  crt.Screen.Send (index & Chr(13))
End Sub

					
					
