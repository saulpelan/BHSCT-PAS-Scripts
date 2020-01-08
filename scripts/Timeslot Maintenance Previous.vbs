# $language = "VBScript"
# $interface = "1.0"



'	Script name: Timeslot Maintenance Previous
'	Author: Saul Pelan
'	Recommended map key: CTRL + -
'
'	┌Description────────────────────────────────────────────────────────────────────┐
'	│ Allows the maintenance user to jump to the Prev timeslot without having to 	│
' 	│ navigate the menu and manually search for timeslots.				│
'	└───────────────────────────────────────────────────────────────────────────────┘
' 
'	┌Intended use───────────────────────────────────────────────────────────────────┐
'	│ To speed up the task of maintaining session timeslots. For example if the	│
'	│ maintenance user is required to revise the location of a timeslot, they just	│
'	│ have to revise the first one and jump to the Prev timeslot using this script	│
'	│ and it will automatically set the cursor location to the same field the user	│
'	│ was amending.									│
'	└───────────────────────────────────────────────────────────────────────────────┘
'
'	┌Instructions───────────────────────────────────────────────────────────────────┐
'	│ Run this script when you are at the Outpatients Maintain Timeslot page in the	│
'	│ Doctor Template or Maintain Doctor Session function. You will be brought to	│
'	│ the Prev timeslot and your cursor remain at the same field for the Prev slot.	│									│
'       └───────────────────────────────────────────────────────────────────────────────┘


scriptName = "Timeslot Maintenance Previous"
errorState = False
sleepMilliseconds = 1
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
			

			If tscmd <> "REVISE" Then
				GoToField(2)
				Send("REVISE")
			End If
			
			If row = 7 Then
				If col <> 20 Then 'If cursor isn't at the first column of the field
					SendSpecial("VT_SELECT") 'Moves cursor to start of field to overwrite any text
				End If
				Send("REVISE") 'Default timeslot command to 'revise'
				SelectPrevTimeslot()
				GoToField(10) 'Go to the Timeslot Patients field to display furhter timeslot details
				GoToField(2) 'Go back to the Timeslot Command field
			Elseif row = 9 Then
				'Check if cursor is at Timeslot Start field
				If col >= 20 And col <= 26 Then
					SelectPrevTimeslot()
					GoToField(3) 'Go back to Timeslot Start field

				'Check if cursor is at Timeslot Stop field
				Elseif col >= 41 And col <= 47 Then
					GoToField(3) 'Navigate to Timeslot Start field so we can select a new one
					SelectPrevTimeslot()
				Else	
					Msg  "Couldn't identify cursor column on row " & row & ". How did you manage that?!"
				End If	
			Elseif row = 16 Then
				
				'Check if cursor is at Timeslot Patients field
				If col >= 22 And col <= 24 Then
					GoToField(3)
					SelectPrevTimeslot()
					GoToField(10)
				
				'Check if cursor is at Report-To Location 
				Elseif col >= 43 And col <= 47 Then
					GoToField(3)
					SelectPrevTimeslot()
					GoToField(11)
				Else
					Msg "Couldn't identify cursor column on row " & row & ". How did you manage that?!" 
				End If
			
			'Check if cursor is anywhere on the terminal scroll list for appointment types
			Elseif row >= 12 And row <= 15 Then
				GoToField(3)
				SelectPrevTimeslot()
				GoToField(6)
				Send("1")
			Else 
				Msg "Invalid cursor position for script."
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
			

			If tscmd <> "REVISE" Then
				GoToField(2)
				Send("REVISE")
			End If
			
			If row = 8 Then
				If col <> 20 Then 'If cursor isn't at the first column of the field
					SendSpecial("VT_SELECT") 'Moves cursor to start of field to overwrite any text
				End If
				Send("REVISE") 'Default timeslot command to 'revise'
				SelectPrevTimeslot()
				GoToField(10) 'Go to the Timeslot Patients field to display furhter timeslot details
				GoToField(2) 'Go back to the Timeslot Command field
			Elseif row = 10 Then
				'Check if cursor is at Timeslot Start field
				If col >= 20 And col <= 26 Then
					SelectPrevTimeslot()
					GoToField(3) 'Go back to Timeslot Start field

				'Check if cursor is at Timeslot Stop field
				Elseif col >= 41 And col <= 47 Then
					GoToField(3) 'Navigate to Timeslot Start field so we can select a new one
					SelectPrevTimeslot()
				Else	
					Msg  "Couldn't identify cursor column on row " & row & ". How did you manage that?!"
				End If	
			Elseif row = 17 Then
				
				'Check if cursor is at Timeslot Patients field
				If col >= 22 And col <= 24 Then
					GoToField(3)
					SelectPrevTimeslot()
					GoToField(10)
				
				'Check if cursor is at Report-To Location 
				Elseif col >= 43 And col <= 47 Then
					GoToField(3)
					SelectPrevTimeslot()
					GoToField(11)
				Else
					Msg "Couldn't identify cursor column on row " & row & ". How did you manage that?!" 
				End If
			
			'Check if cursor is anywhere on the terminal scroll list for appointment types
			Elseif row >= 13 And row <= 16 Then
				GoToField(3)
				SelectPrevTimeslot()
				GoToField(6)
				Send("1")
			Else 
				Msg "Invalid cursor position for script."
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

Sub SelectPrevTimeslot()
	SendSpecial("VT_F14") 'Open superhelp list of timeslots
	SendSpecial("VT_CURSOR_UP") 'Select Prev timeslot
	Send("") 'send return to select timeslot from list
End Sub

Sub Db(ByVal text)
	If test Then
		MsgBox "** Timeslot Maintenance Prev **" & VbCrLf & text
	End If
End Sub

Sub Msg(ByVal text)
	MsgBox "** " & scriptName & " **" & VbCrLf & text
End Sub
