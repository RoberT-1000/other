'STATS GATHERING ---------------------------
'name_of_script = "ACTIONS - REDIRECT DOCS.vbs"
'start_time = timer


'LOADING SCRIPT
'url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
'SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
'req.open "GET", url, FALSE									'Attempts to open the URL
'req.send													'Sends request
'IF req.Status = 200 THEN									'200 means great success
'	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
'	Execute req.responseText								'Executes the script code
'ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
'	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
'			vbCr & _
'			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
'			vbCr & _
'			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
'			vbTab & "- The name of the script you are running." & vbCr &_
'			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
'			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
'			vbTab & vbTab & "responsible for network issues." & vbCr &_
'			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
'			vbCr & _
'			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
'			vbCr &_
'			"URL: " & url
'			StopScript
'END IF

'thirty_days_away = DateAdd("d", 30, date) 
'month_after = DateAdd("m", 1, thirty_days_away)
'redirection_month = DatePart("m", month_after)
'redirection_year = DatePart("yyyy", month_after)
'If len(redirection_month) = 1 then redirection_month = "0" & redirection_month

'Dim case_number, caregiver_case_number, caregiver_name, prorate_yes, prorate_no, child_one, child_two, child_three, child_four, child_five, child_six, cch_amount, cms_amount, ccc_amount, total_amount, original_cp_name
BeginDialog redirection_dialog, 0, 0, 236, 285, "Redirection Script"
'  EditBox 90, 5, 145, 15, case_number
'  EditBox 90, 20, 145, 15, caregiver_case_number
'  EditBox 90, 35, 145, 15, original_cp_name
'  EditBox 90, 50, 145, 15, caregiver_name
'  CheckBox 115, 75, 25, 10, "Yes", prorate_yes
'  CheckBox 140, 75, 20, 10, "No", prorate_no
'  EditBox 10, 105, 220, 15, child_one
'  EditBox 10, 120, 220, 15, child_two
'  EditBox 10, 135, 220, 15, child_three
'  EditBox 10, 150, 220, 15, child_four
'  EditBox 10, 165, 220, 15, child_five
'  EditBox 10, 180, 220, 15, child_six
'  EditBox 5, 210, 70, 15, cch_amount
'  EditBox 80, 210, 70, 15, cms_amount
'  EditBox 155, 210, 70, 15, ccc_amount
'  EditBox 80, 235, 70, 15, total_amount
'  ButtonGroup ButtonPressed
    OkButton 55, 260, 50, 15
    CancelButton 125, 260, 50, 15
'  Text 40, 10, 50, 10, "Case Number:"
'  Text 5, 25, 80, 10, "Caregiver Case Number:"
'  Text 20, 55, 65, 10, "Name Of Caregiver:"
'  Text 5, 75, 105, 10, "Prorate Support Per Each Child?"
'  Text 5, 95, 200, 10, "Child(ren) that are having support redirected (FULL NAME)"
'  Text 10, 200, 180, 10, "Amounts To Be Redirected: (must be xxx.xx format)"
'  Text 30, 225, 20, 10, "CCH"
'  Text 100, 225, 35, 10, "CMS"
'  Text 180, 225, 35, 10, "CCC"
'  Text 95, 250, 40, 10, "TOTAL"
'  Text 25, 40, 60, 10, "Original CP name:"
EndDialog


'Connects to Bluezone
EMConnect ""

'Starts dialog
					Dialog redirection_dialog
     					IF ButtonPressed = 0 THEN StopScript

'goes to correct case
'EMWriteScreen "CAST", 21,18
'Transmit
'EMWriteScreen "D", 3, 29
'Puts case number in from Dialog box
'	EMWriteScreen Left (case_number, 10), 4, 8
'	EMWriteScreen Right (case_number, 2), 4, 19
'Transmit

'________________________________________________________________________________________________________________________________________________________________________________________ NCP NOTICE
'goes to DORD
EMWriteScreen "DORD", 21,18
Transmit
'Adds dord doc
EMWriteScreen "A", 3, 29
'blanks out any DOC ID number that may be entered
EMWriteScreen "        ", 4, 50
EMWriteScreen "       ", 4, 59
EMWriteScreen "F0500", 6, 36
Transmit

'entering user labels
EMSendKey (PF14)
EMWriteScreen "U", 20, 14
Transmit
EMWriteScreen "S", 7, 5
EMWriteScreen "S", 8, 5
EMWriteScreen "S", 9, 5
EMWriteScreen "S", 10, 5
EMWriteScreen "S", 11, 5
EMWriteScreen "S", 12, 5
EMWriteScreen "S", 13, 5
EMWriteScreen "S", 14, 5
EMWriteScreen "S", 15, 5
EMWriteScreen "S", 16, 5
EMWriteScreen "S", 17, 5
EMWriteScreen "S", 18, 5
Transmit

EMWriteScreen redirection_month, 16, 15
EMWriteScreen "/01/", 16, 17
EMWriteScreen redirection_year, 16, 21
Transmit
EMWriteScreen caregiver_name, 16, 15
Transmit
If prorate_yes = checked then
	EMWriteScreen "Y", 16, 15
	Transmit
End If
If prorate_no = checked then
	EMWriteScreen "N", 16, 15
	Transmit
End If
EMWriteScreen child_one, 16, 15
Transmit
EMWriteScreen child_two, 16, 15
Transmit
EMWriteScreen child_three, 16, 15
Transmit
EMWriteScreen child_four, 16, 15
Transmit
EMWriteScreen child_five, 16, 15
Transmit
EMWriteScreen child_six, 16, 15
Transmit
EMWriteScreen cch_amount, 16, 15
Transmit
EMWriteScreen cms_amount, 16, 15
Transmit
EMWriteScreen ccc_amount, 16, 15
Transmit
EMSendKey (PF8)
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen total_amount, 16, 15
Transmit
EMSendKey (PF3)
EMWriteScreen "M", 3, 29
Transmit								'At this point, the notice to the NCP is ready to be printed
'________________________________________________________________________________________________________________________________________________________________________________________ CP NOTICE
'clears DORD screen and adds and completes notice to CP
EMWriteScreen "C", 3, 29
Transmit

EMWriteScreen "A", 3, 29
EMWriteScreen "F0501", 6, 36
Transmit

'entering user labels
EMSendKey (PF14)
EMWriteScreen "U", 20, 14
Transmit
EMWriteScreen "S", 7, 5
EMWriteScreen "S", 8, 5
EMWriteScreen "S", 9, 5
EMWriteScreen "S", 10, 5
EMWriteScreen "S", 11, 5
EMWriteScreen "S", 12, 5
EMWriteScreen "S", 13, 5
EMWriteScreen "S", 14, 5
EMWriteScreen "S", 15, 5
EMWriteScreen "S", 16, 5
EMWriteScreen "S", 17, 5
EMWriteScreen "S", 18, 5
Transmit

EMWriteScreen redirection_month, 16, 15
EMWriteScreen "/01/", 16, 17
EMWriteScreen redirection_year, 16, 21
Transmit
EMWriteScreen caregiver_name, 16, 15
Transmit
If prorate_yes = checked then
	EMWriteScreen "Y", 16, 15
	Transmit
End If
If prorate_no = checked then
	EMWriteScreen "N", 16, 15
	Transmit
End If
EMWriteScreen child_one, 16, 15
Transmit
EMWriteScreen child_two, 16, 15
Transmit
EMWriteScreen child_three, 16, 15
Transmit
EMWriteScreen child_four, 16, 15
Transmit
EMWriteScreen child_five, 16, 15
Transmit
EMWriteScreen child_six, 16, 15
Transmit
EMWriteScreen cch_amount, 16, 15
Transmit
EMWriteScreen cms_amount, 16, 15
Transmit
EMWriteScreen ccc_amount, 16, 15
Transmit
EMSendKey (PF8)
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen total_amount, 16, 15
Transmit
EMSendKey (PF3)
EMWriteScreen "M", 3, 29
Transmit								'At this point, the notice to the CP is ready to be printed

'Enters worklist explaining to start redirection effective for the following month. 
EMWriteScreen "CAWT", 21,18
Transmit
EMSendKey (PF5)
EMWriteScreen "A", 3, 30
EMWriteScreen "FREE", 4, 37
EMWriteScreen "The redirection should be effective the 1st of next month", 10, 4 
EMWriteScreen "30", 17, 52
Transmit 

'________________________________________________________________________________________________________________________________________________________________________________________ CAREGIVER NOTICE

'goes to caregiver case
EMWriteScreen "CAST", 21,18
Transmit
EMWriteScreen "D", 3, 29
'Puts caregiver case number in from Dialog box
	EMWriteScreen Left (caregiver_case_number, 10), 4, 8
	EMWriteScreen Right (caregiver_case_number, 2), 4, 19
Transmit
'goes to DORD
EMWriteScreen "DORD", 21,18
Transmit
'Adds dord doc
EMWriteScreen "A", 3, 29
'blanks out any DOC ID number that may be entered
EMWriteScreen "        ", 4, 50
EMWriteScreen "       ", 4, 59
EMWriteScreen "F0502", 6, 36
Transmit

'entering user labels
EMSendKey (PF14)
EMWriteScreen "U", 20, 14
Transmit
EMWriteScreen "S", 7, 5
EMWriteScreen "S", 8, 5
EMWriteScreen "S", 9, 5
EMWriteScreen "S", 10, 5
EMWriteScreen "S", 11, 5
EMWriteScreen "S", 12, 5
EMWriteScreen "S", 13, 5
EMWriteScreen "S", 14, 5
EMWriteScreen "S", 15, 5
EMWriteScreen "S", 16, 5
EMWriteScreen "S", 17, 5
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen redirection_month, 16, 15
EMWriteScreen "/01/", 16, 17
EMWriteScreen redirection_year, 16, 21
Transmit
EMWriteScreen original_cp_name, 16, 15                                  
Transmit
EMWriteScreen child_one, 16, 15
Transmit
EMWriteScreen child_two, 16, 15
Transmit
EMWriteScreen child_three, 16, 15
Transmit
EMWriteScreen child_four, 16, 15
Transmit
EMWriteScreen child_five, 16, 15
Transmit
EMWriteScreen child_six, 16, 15
Transmit
EMWriteScreen cch_amount, 16, 15
Transmit
EMWriteScreen cms_amount, 16, 15
Transmit
EMWriteScreen ccc_amount, 16, 15
Transmit
EMWriteScreen total_amount, 16, 15
Transmit
EMSendKey (PF3)
EMWriteScreen "M", 3, 29
Transmit								'At this point, the notice to the caregiver is ready to be printed

'Enters worklist explaining to start redirection effective for the following month. 
EMWriteScreen "CAWT", 21,18
Transmit
EMSendKey (PF5)
EMWriteScreen "A", 3, 30
EMWriteScreen "FREE", 4, 37
EMWriteScreen "The redirection should be effective the 1st of next month", 10, 4 
EMWriteScreen "30", 17, 52
Transmit 

script_end_procedure("DORD docs are created but incomplete.  Please modify to select appropriate legal headings.")
