'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MNSURE NOTES - MNSURE DOCS RECEIVED.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

BeginDialog client_contact_dialog, 0, 0, 386, 200, "MNSure Client contact"
  ComboBox 50, 5, 60, 15, "Phone call"+chr(9)+"Voicemail"+chr(9)+"Email"+chr(9)+"Office visit"+chr(9)+"Letter", contact_type
  DropListBox 115, 5, 45, 10, "from"+chr(9)+"to", contact_direction
  ComboBox 165, 5, 85, 15, "client"+chr(9)+"AREP"+chr(9)+"Non-AREP"+chr(9)+"SWKR", who_contacted
  EditBox 280, 5, 100, 15, regarding
  EditBox 80, 25, 65, 15, phone_number
  EditBox 290, 25, 85, 15, when_contact_was_made
  EditBox 290, 45, 85, 15, when_contact_was_returned
  EditBox 70, 70, 310, 15, contact_reason
  EditBox 55, 90, 325, 15, actions_taken
  EditBox 75, 125, 300, 15, verifs_needed
  EditBox 60, 145, 315, 15, case_status
  EditBox 185, 180, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 275, 180, 50, 15
    CancelButton 330, 180, 50, 15
  Text 5, 10, 45, 10, "Contact type:"
  Text 260, 10, 15, 10, "Re:"
  Text 25, 30, 50, 10, "Phone number: "
  Text 180, 30, 95, 10, "Date/Time of Client Contact"
  Text 140, 50, 135, 10, "Date/Time of Worker Contact (if different)"
  Text 5, 75, 65, 10, "Reason for contact:"
  Text 5, 95, 50, 10, "Actions taken: "
  GroupBox 5, 110, 375, 60, "Other information"
  Text 10, 130, 60, 10, "Evidence needed: "
  Text 10, 150, 45, 10, "Case status: "
  Text 110, 185, 70, 10, "Sign your case note: "
EndDialog

DO
	err_msg = "" 
	DIALOG client_contact_dialog
		IF ButtonPressed = 0 THEN stopscript
		IF contact_type = "" THEN err_msg = err_msg & vbCr & "* 'Contact type' is blank. Please indicate the origination of the client contact."
		IF who_contacted = "" THEN err_msg = err_msg & vbCr & "* Please indicate with whom you contacted."
		IF regarding = "" THEN err_msg = err_msg & vbCr & "* 'RE:' is blank. Please enter the subject of the client contact."
		IF when_contact_was_made = "" THEN err_msg = err_msg & vbCr & "* Please enter a date/time for the client contact."
		IF contact_reason = "" THEN err_msg = err_msg & vbCr & "* 'Reason for contact' is blank. Please enter the reason for the contact."
		IF actions_taken = "" THEN err_msg = err_msg & vbCr & "* Please discuss the actions taken."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""
		
'Now it creates a Word document to store all of the active claims.
Set objWord = CreateObject("Word.Application")
objWord.Visible = true
set objDoc = objWord.Documents.add()
Set objSelection = objWord.Selection
objSelection.ParagraphFormat.LineSpacing = 12
objSelection.Range.ParagraphFormat.SpaceAfter = 0

objSelection.TypeText "-- CLIENT CONTACT: " & contact_type & " " & contact_direction & " " & who_contacted & ", RE: " & regarding & " --"
objSelection.TypeParagraph()
objSelection.TypeText "* Client Contact Date/Time: " & when_contact_was_made
objSelection.TypeParagraph()
IF when_contact_was_returned <> "" THEN 
	objSelection.TypeText "* Worker Contact Date/Time: " & when_contact_was_returned
	objSelection.TypeParagraph()
END IF
IF phone_number <> "" THEN 
	objSelection.TypeText "* Phone Number: " & phone_number
	objSelection.TypeParagraph()
END IF
objSelection.TypeText "* Reason for Contact: " & contact_reason
objSelection.TypeParagraph()
objSelection.TypeText "* Actions Taken: " & actions_taken
objSelection.TypeParagraph()
IF verifs_needed <> "" THEN 
	objSelection.TypeText "* Evidence Needed: " & verifs_needed
	objSelection.TypeParagraph()
END IF
IF case_status <> "" THEN 
	objSelection.TypeText "* Case Status: " & case_status
	objSelection.TypeParagraph()
END IF
objSelection.TypeText "***"
objSelection.TypeParagraph()
objSelection.TypeText worker_signature
