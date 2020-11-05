Class ClinicGroup

	Private ioh

	
	' Initialises the ClinicGroup object as soon as it is created with a link to the
	' IO class.

	Private Sub Class_Initialize()
		Include "library/IO.vbs"
		Set IOHandler = New IO
  	End Sub


	' Sets the linked IO object to <io1>.

	Private Property Set IOHandler(io1)
		Set ioh = io1
	End Property


	' Returns the linked IO object.

	Public Property Get IOHandler()
		Set IOHandler = ioh
	End Property


	' Attempts to perform a REVISE command on the clinic group of name <clinicGroup>
	' and navigate to the list of clinic codes.
	'
	' Returns True if the script was able to complete the action or False if not

	Function Revise(ByVal clinicGroup)
		Revise = False
		If IOHandler.GetTitle(True) = "OP Clinic Group Master File" And Not IsEditing() Then
			IOHandler.ResetForm
			IOHandler.SendText "REVISE", True
			If IOHandler.SendCarefully(clinicGroup) Then
				IOHandler.SendText "", True
				screen1 = crt.Screen.Get2(1, 1, crt.Screen.Rows, crt.Screen.Columns)
				IOHandler.GoToField "5"
				If IOHandler.WaitForScreenUpdate(screen1, 1) Then
					Revise = IsEditing()
				End If
			End If
		End If
	End Function


	' Gets the currently selected clinic code.
	'
	' Returns selected clinic code as String

	Function GetSelectedClinicCode()
		If IsEditing Then
			row = crt.Screen.CurrentRow
			If row >= 6 And Row <= 15 Then
				GetSelectedClinicCode = Trim(crt.Screen.Get(row, 19, row, 26))
			End If
		End If
	End Function


	' Attempts to navigate to the clinic at index <index>. <index> should be a 
	' string representation of an integer.
	'
	' Returns True if it was possible to navigate to the chosen index or False
	' if the attempt failed. 

	Function SelectIndex(ByVal index)
		If IsEditing Then
			IOHandler.GoToFieldOccurrence "6", index
		End If
	End Function
	

	' Gets the index number if the currently selected clinic code.
	'
	' Returns an integer representing the index of the selected record

	Function GetSelectedIndex()
		If IsEditing Then
			row = crt.Screen.CurrentRow
			If row >= 6 And Row <= 15 Then
				GetSelectedIndex = Trim(crt.Screen.Get(row, 14, row, 16))
			End If
		End If
	End Function


	' Checks if a Clinic Group is currently being edited and its entries are
	' on display.
	'
	' Returns True if a Clinic Group is being edited and its entries are on display
	' or False if not

	Function IsEditing()
		IsEditing = IOHandler.GetTitle(True) = "OP Clinic Group Master File" And _
			IOHandler.GetSubtitle() = "Outpatients Enter Clinics"
	End Function


	' Gets the Clinic Group that is currently being edited if there is one.
	'
	' Returns the name of the currently selected Clinic Group
	Function GetCurrentGroup()
		If IsEditing() Then
			GetCurrentGroup = Trim(crt.Screen.Get(4, 18, 4, 25))
		End If
	End Function


	' Navigates to the "Enter?" prompt and enters "Yes" if <save> is True, or "No"
	' if <save> is false.

	Sub Enter(ByVal save)
		If IsEditing() Then
			IOHandler.GoToField "7"
			If save Then
				IOHandler.SendText "YES", True
			Else
				IOHandler.SendText "NO", True
			End If
		End If
	End Sub

End Class
