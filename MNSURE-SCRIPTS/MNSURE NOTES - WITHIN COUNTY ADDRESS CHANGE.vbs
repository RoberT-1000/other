'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MNSURE NOTES - ADDR CHANGE WITHIN COUNTY.vbs"
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

BeginDialog mnsure_note_dlg, 0, 0, 306, 345, "Within Address Change"
  EditBox 105, 10, 85, 15, mnsure_case_number
  ComboBox 130, 30, 75, 15, ""+chr(9)+"In Person"+chr(9)+"Client Call"+chr(9)+"Voicemail"+chr(9)+"Returned Mail", addr_change_source
  EditBox 15, 70, 120, 15, new_addr_line1
  EditBox 15, 90, 120, 15, addr_line2
  EditBox 15, 110, 50, 15, new_addr_city
  EditBox 70, 110, 25, 15, new_addr_state
  EditBox 100, 110, 35, 15, new_addr_zip
  EditBox 170, 70, 120, 15, old_addr_line1
  EditBox 170, 90, 120, 15, old_addr_line2
  EditBox 170, 110, 50, 15, old_addr_city
  EditBox 225, 110, 25, 15, old_addr_state
  EditBox 255, 110, 35, 15, old_addr_zip
  CheckBox 10, 145, 290, 10, "Check HERE if the client reports being homeless.", homeless_check
  CheckBox 10, 160, 290, 10, "Check HERE if the new address is a mailing address.", mailing_addr_check
  CheckBox 10, 180, 180, 10, "Check HERE is you updated the tracking database.", updated_tracking_database_check
  CheckBox 10, 195, 280, 10, "Check HERE if you have already updated the Case/Person evidence.", udpated_case_person_evidence_check
  EditBox 75, 215, 65, 15, cl_move_date
  EditBox 135, 240, 60, 15, case_number
  CheckBox 205, 245, 75, 10, "Case Note in MAXIS", case_note_in_maxis_check
  EditBox 60, 260, 240, 15, other_notes
  EditBox 80, 280, 220, 15, evidence_needed
  EditBox 75, 315, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 200, 315, 50, 15
    CancelButton 250, 315, 50, 15
  Text 10, 15, 85, 10, "Integrated Case Number"
  Text 10, 35, 120, 10, "How was address change reported?"
  GroupBox 5, 55, 140, 80, "New Address"
  GroupBox 160, 55, 140, 80, "Old Address"
  Text 10, 220, 60, 10, "Client Move Date:"
  Text 10, 245, 120, 10, "MAXIS Case Number (blank if none)"
  Text 10, 265, 45, 10, "Other Notes:"
  Text 10, 285, 65, 10, "Evidence Needed:"
  Text 10, 320, 60, 10, "Worker Signature"
EndDialog

new_addr_state = "MN"
old_addr_state = "MN"

DO
	err_msg = ""
	DIALOG mnsure_note_dlg
		IF ButtonPressed = 0 THEN stopscript
		IF mnsure_case_number = "" THEN err_msg = err_msg & vbCr & "* Please enter the MNSure Integrated Case Number."
		IF addr_change_source = "" THEN err_msg = err_msg & vbCr & "* Please enter or select the source of the change report."
		IF (new_addr_line1 = "" OR new_addr_city = "" OR new_addr_state = "") AND homeless_check = 0 THEN err_msg = err_msg & vbCr & "* Please enter the new address."
		IF cl_move_date = "" THEN err_msg = err_msg & vbCr & "* Please enter the client move date."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'Now it creates a Word document
Set objWord = CreateObject("Word.Application")
objWord.Visible = true
set objDoc = objWord.Documents.add()
Set objSelection = objWord.Selection
objSelection.ParagraphFormat.LineSpacing = 12
objSelection.Range.ParagraphFormat.SpaceAfter = 0

objSelection.TypeText "-- Address Change Within County --"
objSelection.TypeParagraph()
objSelection.TypeText "* Address changed reported via: " & addr_change_source
objSelection.TypeParagraph()
IF homeless_check = 1 THEN 
	objSelection.TypeText "* Client reports being homeless."
	objSelection.TypeParagraph()
END IF
IF new_addr_line1 <> "" THEN 
	objSelection.TypeText "* New Address..."
	objSelection.TypeParagraph()
	objSelection.TypeText "     " & new_addr_line1
	objSelection.TypeParagraph()
	IF new_addr_line2 <> "" THEN 
		objSelection.TypeText "     " & new_addr_line2
		objSelection.TypeParagraph()
	END IF
	objSelection.TypeText "     " & new_addr_city & ", " & new_addr_state & " " & new_addr_zip
	objSelection.TypeParagraph()
	IF mailing_addr_check = 1 THEN 
		objSelection.TypeText "* Client reports this is a mailing address only."
		objSelection.TypeParagraph()
	END IF
END IF
IF old_addr_line1 <> "" THEN 
	objSelection.TypeText "* Old Address..."
	objSelection.TypeParagraph()
	objSelection.TypeText "     " & old_addr_line1
	objSelection.TypeParagraph()
	IF old_addr_line2 <> "" THEN 
		objSelection.TypeText "     " & old_addr_line2
		objSelection.TypeParagraph()
	END IF
	objSelection.TypeText "     " & old_addr_city & ", " & old_addr_state & " " & old_addr_zip
	objSelection.TypeParagraph()
END IF
IF updated_tracking_database_check = 1 THEN 
	objSelection.TypeText "* Updated MNSure tracking database"
	objSelection.TypeParagraph()
END IF
IF udpated_case_person_evidence_check = 1 THEN 
	objSelection.TypeText "* Updated Case & Person Evidence"
	objSelection.TypeParagraph()
END IF
IF maxis_case_number <> "" THEN 
	objSelection.TypeText "* MAXIS Case Number: " & maxis_case_number
	objSelection.TypeParagraph()
END IF
IF other_notes <> "" THEN 
	objSelection.TypeText "* Other Notes: " & other_notes
	objSelection.TypeParagraph()
END IF
IF evidence_needed <> "" THEN 
	objSelection.TypeText "* Evidence Needed: " & evidence_needed
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
		CALL write_variable_in_case_note("*** ADDR CHANGE REPORTED THROUGH MNSURE ***")
		CALL write_bullet_and_variable_in_case_note("Integrated Case Number", mnsure_case_number)
		CALL write_bullet_and_variable_in_case_note("Reported Move Date", cl_move_date)
		IF other_notes <> "" THEN CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
		IF evidence_needed <> "" THEN CALL write_bullet_and_variable_in_case_note("Verifs Needed", evidence_needed)
		IF updated_tracking_database_check = 1 THEN CALL write_variable_in_case_note("* Updated MNSure Tracking Database.")
		CALL write_variable_in_case_note("---")
		CALL write_variable_in_case_note(worker_signature)
	
		CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
		CALL create_MAXIS_friendly_date(date, 0, 5, 18)
		EMWriteScreen "CLIENT REPORTS MOVE EFFECTIVE " & cl_move_date, 9, 3
		EMWriteScreen "CLIENT REPORTS NOW LIVING IN " & new_addr_city, 10, 3
		EMWriteScreen "SEE INTEGRATED CASE " & mnsure_case_number & " FOR MORE INFORMATION", 11, 3
		TRANSMIT
		PF3
	END IF
	
END IF
