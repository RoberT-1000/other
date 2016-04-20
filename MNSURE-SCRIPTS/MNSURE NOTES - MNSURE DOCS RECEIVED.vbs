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

BeginDialog docs_received_dialog, 0, 0, 466, 165, "Docs received"
  EditBox 95, 5, 90, 15, mnsure_case_number
  EditBox 60, 25, 215, 15, docs_received
  EditBox 60, 45, 80, 15, received_date
  EditBox 75, 65, 385, 15, verif_notes
  EditBox 60, 85, 400, 15, actions_taken
  EditBox 135, 105, 325, 15, docs_needed
  EditBox 130, 125, 55, 15, case_number
  CheckBox 195, 130, 160, 10, "Check HERE to Case Note and TIKL in MAXIS", case_note_in_maxis_check
  EditBox 70, 145, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 355, 5, 50, 15
    CancelButton 410, 5, 50, 15
  Text 5, 10, 85, 10, "MNSure Case Number:"
  Text 5, 30, 55, 10, "Docs received:"
  Text 5, 50, 55, 10, "Date Received:"
  Text 280, 30, 190, 10, "Note: just list the docs here. This is the title of your note."
  Text 5, 70, 70, 10, "Notes on your docs:"
  Text 5, 90, 50, 10, "Actions taken: "
  Text 5, 110, 130, 10, "Evidence still needed (if applicable):"
  Text 5, 130, 120, 10, "MAXIS Case Number (blank if none):"
  Text 5, 150, 60, 10, "Worker signature:"
EndDialog


DO
	err_msg = ""
	DIALOG docs_received_dialog
		IF ButtonPressed = 0 THEN stopscript
			IF mnsure_case_number = "" THEN err_msg = err_msg & vbCr & "* Please enter the MNSure Case Number."
			IF docs_received = "" THEN err_msg = err_msg & vbCr & "* Please enter the document(s) received."
			IF received_date = "" THEN err_msg = err_msg & vbCr & "* Please enter the date received."
			IF actions_taken = "" THEN err_msg = err_msg & vbCr & "* Please specify the actions taken."
			IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'Now it creates a Word document
Set objWord = CreateObject("Word.Application")
objWord.Visible = true
set objDoc = objWord.Documents.add()
Set objSelection = objWord.Selection

objSelection.TypeText "-- EVIDENCE RECEIVED --"
objSelection.TypeParagraph()
objSelection.TypeText "* Document(s) Received: " & docs_received
objSelection.TypeParagraph()
objSelection.TypeText "* Date Received: " & received_date
objSelection.TypeParagraph()
objSelection.ParagraphFormat.LineSpacing = 12
objSelection.Range.ParagraphFormat.SpaceAfter = 0

IF verif_notes <> "" THEN 
	objSelection.TypeText "* Notes on Doc(s): " & verif_notes
	objSelection.TypeParagraph()
END IF
objSelection.TypeText "* Actions Taken: " & actions_taken
objSelection.TypeParagraph()
IF case_number <> "" THEN 
	objSelection.TypeText "* MAXIS Case Number: " & case_number
	objSelection.TypeParagraph()
	IF case_note_in_maxis_check = 1 THEN 
		objSelection.TypeText "*      Case noted in MAXIS."
		objSelection.TypeParagraph()
	END IF
END IF
IF docs_needed <> "" THEN 
	objSelection.TypeText "* Evidence Needed: " & docs_needed
	objSelection.TypeParagraph()
END IF
objSelection.TypeText "***"
objSelection.TypeParagraph()
objSelection.TypeText worker_signature

IF case_note_in_maxis_check = 1 AND case_number <> "" THEN 

	EMConnect ""

	CALL check_for_MAXIS(False)
	CALL navigate_to_MAXIS_screen("CASE", "NOTE")
	EMReadScreen privileged_look_up, 10, 24, 14
	EMReadScreen invalid_case_number, 7, 24, 2
	IF privileged_look_up = "PRIVILEGED" THEN script_end_procedure("MAXIS case number " & case_number & " is privileged. The script cannot case note or TIKL on this case. The script will now end.")
	IF invalid_case_number = "INVALID" THEN script_end_procedure("MAXIS case number " & case_number & " is invalid. The script will now end.")
	PF9
	EMReadScreen error_msg, 9, 24, 12
	IF error_msg = "READ ONLY" THEN 
		MsgBox("Cannot case note on MAXIS case " & case_number & ". The script will now go straight to TIKL.")
	ELSE
		CALL write_variable_in_case_note("*** DOCS RECEIVED FOR MNSURE ***")
		CALL write_bullet_and_variable_in_case_note("MNSure Case Number", mnsure_case_number)
		CALL write_bullet_and_variable_in_case_note("Docs Received", docs_received)
		CALL write_bullet_and_variable_in_case_note("Received Date", received_date)
		IF verif_notes <> "" THEN CALL write_bullet_and_variable_in_case_note("Notes on Docs", verif_notes)
		CALL write_bullet_and_variable_in_case_note("Actions Taken", actions_taken)
		IF docs_needed <> "" THEN CALL write_bullet_and_variable_in_case_note("Verifs Needed", docs_needed)
		CALL write_variable_in_case_note("---")
		CALL write_variable_in_case_note(worker_signature)
	END IF
	
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	CALL create_MAXIS_friendly_date(date, 0, 5, 18)
	EMWriteScreen "MNSURE DOCS RECEIVED ON " & received_date, 9, 3
	EMWriteScreen "SEE MNSURE CASE " & mnsure_case_number & " FOR MORE INFORMATION", 10, 3
	TRANSMIT
	PF3
END IF

