# $language = "VBScript"
# $interface = "1.0"

Dim hospital
Dim excelApp
Dim workbook
Dim worksheet
Dim tab
Dim fnColumn
Dim snColumn
Dim dobColumn
Dim curRow

Sub Main()

	set excelApp = GetObject(,"Excel.Application")
	set workbook = excelApp.Workbooks(1)

	fnColumn = "F"
	snColumn = "G"
	dobColumn = "I"
			
	If InStr(workbook.Name, "Mater") > 0 Then
		hospital = "Mater"
	Elseif InStr(workbook.Name, "RGH") > 0 Then
		hospital = "Royal"
	Elseif InStr(workbook.Name, "BCH") > 0 Then
		hospital = "Belfast City"
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
					
	MsgBox "Working on Data Quality Report: " & workbook.Name & vbCrLf & "Working on PAS session: " & tab.Caption & vbCrLf & "Report: " & workbook.ActiveSheet.Name
	
	MarkData ActiveRow, "Good"
	
End Sub

Sub GoToField(ByVal index)
        crt.Screen.SendSpecial ("VT_F9")
        crt.Screen.Send (index & Chr(13))
End Sub

Function GetTotalRows()
	i = 4
	set worksheet = workbook.ActiveSheet
	Do While i < 100 AND NOT (IsEmpty(worksheet.Range(fnColumn & i)) AND IsEmpty(worksheet.Range(snColumn & i)) AND IsEmpty(worksheet.Range(dobColumn & i)))
		i = i + 1
	Loop
	GetTotalRows = i - 4
End Function

Function GetActiveRow()
	For i = 4 To GetTotalRows()
		set rng = workbook.ActiveSheet.Range(fnColumn & i & "," & snColumn & i & "," & dobColumn & i)
		If rng.Style = "Normal" Then
			GetActiveRow = i
			Exit For
		End If
	Next
End Function

Sub MarkData(ByVal row, Byval style)
	set rng = workbook.ActiveSheet.Range(fnColumn & row & "," & snColumn & row & "," & dobColumn & row)
	rng.Style = style
End Sub
