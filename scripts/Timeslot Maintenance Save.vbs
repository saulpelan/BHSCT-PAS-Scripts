# $language = "VBScript"
# $interface = "1.0"



'	Script name: Timeslot Maintenance Save
'	Author: Saul Pelan
'	Recommended map key: CTRL + ]
'
'	┌Description────────────────────────────────────────────────────────────────────┐
'	│ Allows the maintenance user to save their changes when maintaining a timeslot.│
'	│										│
'	│ If the user is ADDing a timeslot the minimum criteria is to have the Timeslot	│
'	│ Start and Timeslot Stop. If no appointment type is entered it will use super-	│
'	│ help menu. If there is only one valid appointment type for the clinic it will │
'	│ select this automatically. Otherwise it will wait for the user to select an	│
'	│ appointment type in the superhelp list, and then complete the adding of the	│
'	│ slot, automatically filling in the next Timeslot Start time with the End time	│
'	│ of the previous slot - clever eh?						│
'	│										│
'	│ If the user is REVISING an existing timeslot it will save the timeslot the	│
'	│ way it currently is and then bring the user back to the same field in the same│
'	│ time slot. This allows the user to use the "Timeslot Maintenance Prev/Next"	│
'	│ scripts to jump to the next or previous timeslot and quickly change the same	│
'	│ field you were on previously - allowing you to quickly make changes to multi-	│
'	│ ple time slots.								│
'	└───────────────────────────────────────────────────────────────────────────────┘
' 
'	┌Instructions───────────────────────────────────────────────────────────────────┐
'	│ Run this script when you are at the Outpatients Maintain Timeslot page in the	│
'	│ Doctor Template or Maintain Doctor Session function. The current timeslot will│
'	│ be saved when revising existing slots or adding new slots. Details above.	│									│
'       └───────────────────────────────────────────────────────────────────────────────┘


scriptName = "Timeslot Maintenance Save"
errorState = False
sleepMilliseconds = 100
test = False

Sub Main()
	If Not crt.Session.Connected Then
		Exit Sub
	End If
	title = GetFunctionTitle()
	If title = "M a i n t a i n   D o c t o r   S e s s i o n" Then
		If InStr(GetFunctionSubtitle, "Maintain Timeslot") Then
	
			tscmd = GetText(7, 20, 7, 28)
			row = crt.Screen.CurrentRow
			col = crt.Screen.CurrentColumn
			
			'Check if we are adding a new timeslot
			If tscmd = "ADD" Then
				
				stopTime = GetText(9, 41, 9, 47)

				'Check if Timeslot Stop field is empty
				If stopTime = "" Then
					Msg "Must enter stop time when command is ADD"

				'Check if Appt Type field is empty
				Elseif GetText(12, 16, 12, 18) = "" Then
					Send("")
					GoToField(6)
					Send("1")
					SendSpecial("VT_F14") 'Display Appointment Type superhelp
					ctime = Timer
					
					'Wait for Appointment Type superhelp to appear
					Do While GetText(4, 55, 4, 71) <> "Appointment types"

						'Check if more than 5 seconds has elapsed since trying to display superhelp
						If Timer - ctime > 5 Then
							Msg "Error when trying to use Appointment Type superhelp. Script stopped."
							Exit Sub
						End If
					Loop

					If GetText(4, 55, 4, 71) = "Appointment types" Then

						'Check if second allowed appointment type is blank
						If GetText(7, 55, 7, 57) = "" Then
							Send("") 'Select the only available appointment type in the Superhelp list
						Else
							'Wait until Appt Type superhelp disappears (ie when an appt type is selected)
							Do While GetText(4, 55, 4, 71) = "Appointment types"
								If Timer - ctime > 20 Then
									Msg "Hey wake up! You took too long to select an appointment type. Script stopped. Get back to work!"
									Exit Sub
								End If
							Loop
						End If
					Else
						Msg "I don't know what you've done. Script stopped."
						Exit Sub
					End If
				End If
				stopTime = GetText(9, 41, 9, 47)
				SaveTimeslot()
				GoToField(3)
				SendSpecial("VT_F14")
				SendNoReturn(stopTime)
			'Check if we are revising an existing timeslot
			Elseif tscmd = "REVISE" Then

				'Check if cursor is in the Timeslot Command field
				If row = 7 Then
					SaveTimeslot()
					GoToField(2) 'Go back to the Timeslot Command field
				Elseif row = 9 Then

					'Check if cursor is at Timeslot Start field
					If col >= 20 And col <= 26 Then
						SaveTimeslot()
						GoToField(3) 'Go back to Timeslot Start field	
	
					'Check if cursor is at Timeslot Stop field
					Elseif col >= 41 And col <= 47 Then
						SaveTimeslot()
						GoToField(5) 'Go back to Timeslot Stop field
					Else	
						Msg  "Couldn't identify cursor column on row " & row & ". How did you manage that?!"
					End If	
				Elseif row = 16 Then
					
					'Check if cursor is at Timeslot Patients field
					If col >= 22 And col <= 24 Then
						SaveTimeslot()
						GoToField(10) 'Go back to Timeslot Patients field
					
					'Check if cursor is at Report-To Location 
					Elseif col >= 43 And col <= 47 Then
						SaveTimeslot()
						GoToField(11)
					Else
						Msg "Couldn't identify cursor column on row " & row & ". How did you manage that?!" 
					End If
				
				'Check if cursor is anywhere on the terminal scroll list for appointment types
				Elseif row >= 12 And row <= 15 Then
					SaveTimeslot()
					GoToField(6)
					Send("1")
				Else 
					Msg "Invalid cursor position for script."
				End If
			End If
		Else
			Msg "Not yet in Timeslot Maintenance Page of " & title
			Exit Sub
		End If
	Elseif title = "D o c t o r   T e m p l a t e" Then
		If InStr(GetFunctionSubtitle, "Maintain Timeslot") Then
	
			tscmd = GetText(8, 20, 8, 28)
			row = crt.Screen.CurrentRow
			col = crt.Screen.CurrentColumn
			
			'Check if we are adding a new timeslot
			If tscmd = "ADD" Then
				
				'Check if Timeslot Stop field is empty
				If GetText(10, 41, 10, 47) = "" Then
					Msg "Must enter stop time when command is ADD"

				'Check if Appt Type field is empty
				Elseif GetText(13, 16, 13, 18) = "" Then
					Send("")
					GoToField(6)
					Send("1")
					SendSpecial("VT_F14") 'Display Appointment Type superhelp
					ctime = Timer
					
					'Wait for Appointment Type superhelp to appear
					Do While GetText(4, 55, 4, 71) <> "Appointment types"

						'Check if more than 5 seconds has elapsed since trying to display superhelp
						If Timer - ctime > 5 Then
							Msg "Error when trying to use Appointment Type superhelp. Script stopped."
							Exit Sub
						End If
					Loop

					If GetText(4, 55, 4, 71) = "Appointment types" Then

						'Check if second allowed appointment type is blank
						If GetText(7, 55, 7, 57) = "" Then
							Send("") 'Select the only available appointment type in the Superhelp list
						Else
							'Wait until Appt Type superhelp disappears (ie when an appt type is selected)
							Do While GetText(4, 55, 4, 71) = "Appointment types"
								If Timer - ctime > 20 Then
									Msg "Hey wake up! You took too long to select an appointment type. Script stopped. Get back to work!"
									Exit Sub
								End If
							Loop
						End If
					Else
						Msg "I don't know what you've done. Script stopped."
						Exit Sub
					End If
				End If
				SaveTimeslot()
				GoToField(3)
				SendSpecial("VT_F14")

			'Check if we are revising an existing timeslot
			Elseif tscmd = "REVISE" Then

				'Check if cursor is in the Timeslot Command field
				If row = 8 Then
					SaveTimeslot()
					GoToField(2) 'Go back to the Timeslot Command field
				Elseif row = 10 Then

					'Check if cursor is at Timeslot Start field
					If col >= 20 And col <= 26 Then
						SaveTimeslot()
						GoToField(3) 'Go back to Timeslot Start field	
	
					'Check if cursor is at Timeslot Stop field
					Elseif col >= 41 And col <= 47 Then
						SaveTimeslot()
						GoToField(5) 'Go back to Timeslot Stop field
					Else	
						Msg  "Couldn't identify cursor column on row " & row & ". How did you manage that?!"
					End If	
				Elseif row = 17 Then
					
					'Check if cursor is at Timeslot Patients field
					If col >= 22 And col <= 24 Then
						SaveTimeslot()
						GoToField(10) 'Go back to Timeslot Patients field
					
					'Check if cursor is at Report-To Location 
					Elseif col >= 43 And col <= 47 Then
						SaveTimeslot()
						GoToField(11)
					Else
						Msg "Couldn't identify cursor column on row " & row & ". How did you manage that?!" 
					End If
				
				'Check if cursor is anywhere on the terminal scroll list for appointment types
				Elseif row >= 13 And row <= 16 Then
					SaveTimeslot()
					GoToField(6)
					Send("1")
				Else 
					Msg "Invalid cursor position for script."
				End If
			End If
		Else
			Msg "Not yet in Timeslot Maintenance Page of " & title
			Exit Sub
		End If
	Else
		Msg "Not in Maintain Session or Doctor Template function."
		Exit Sub
	End If
		
End Sub

Sub GoToField(ByVal index)
        SendSpecial("VT_F9")
        Send(index)
End Sub

Sub GoToFieldOcc(ByVal index)
        SendSpecial("VT_F9")
        Send(index)
	Send("1")	
End Sub

Function IsMainMenuVisible()
	IsMainMenuVisible = crt.Screen.Get(1,41,1,49) = "/dev/pts/"
End Function

Function SendSpecial(ByVal text)
	If NOT errorState Then
		crt.Screen.SendSpecial text
		If InStr(UCase(GetStatusText), "ERROR") Then
			errorState = True
		End If
	End If
End Function

Function Send(ByVal text)
	If NOT errorState Then
		crt.Screen.Send text & Chr(13)
		If InStr(UCase(GetStatusText()), "ERROR") Then
			errorState = True
		End If
	End If
End Function

Function SendNoReturn(ByVal text)
	If NOT errorState Then
		crt.Screen.Send text
		If InStr(UCase(GetStatusText()), "ERROR") Then
			errorState = True
		End If
	End If
End Function

Function AttemptSend(ByVal text)
	If NOT errorState Then
		crt.Screen.Send text & Chr(13)
		If InStr(UCase(GetStatusText()), "ERROR") Then
			AttemptSend = False
		Else
			AttemptSend = True
		End If
	End If
End Function

Function GetText(ByVal row1, ByVal col1, ByVal row2, ByVal col2)
	GetText = Trim(crt.Screen.Get(row1, col1, row2, col2))
End Function

Function GetStatusText()
	GetStatusText = GetText(24,1,24,80)
End Function

Function GetFunctionTitle()
	GetFunctionTitle = GetText(1,1,1,80)
End Function

Function GetFunctionSubtitle()
	GetFunctionSubtitle = GetText(2,1,2,41)
End Function

Sub SaveTimeslot()
	Send("") 'Save current field
	GoToField(11) 'Go to Rpt-To Location field (last selectable field in Go to Field selector)
	Send("") 'Send return to proceed to the Enter field
	Send("Y") 'Send YES to save timeslot changes.
End Sub

Sub Db(ByVal text)
	If test Then
		MsgBox "** " & scriptName & " **" & VbCrLf & text
	End If
End Sub

Sub Msg(ByVal text)
	MsgBox "** " & scriptName & " **" & VbCrLf & text
End Sub
