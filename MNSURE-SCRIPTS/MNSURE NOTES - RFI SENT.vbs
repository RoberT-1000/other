'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MNSURE NOTES - RFI SENT.vbs"
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

BeginDialog evidence_needed_dialog, 0, 0, 351, 185, "Evidence Needed"
  EditBox 90, 5, 70, 15, mnsure_case_number
  EditBox 250, 5, 60, 15, verif_due_date
  EditBox 90, 45, 255, 15, hh_comp
  EditBox 40, 65, 305, 15, income
  EditBox 50, 85, 295, 15, residence
  EditBox 65, 105, 280, 15, other_proofs
  CheckBox 5, 135, 175, 10, "Sent form to AREP?", sent_arep_checkbox
  CheckBox 5, 150, 130, 10, "Check here to interact with MAXIS.", MAXIS_check
  Text 15, 170, 75, 10, "MAXIS Case Number:"
  EditBox 100, 165, 65, 15, case_number
  EditBox 285, 140, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    PushButton 315, 10, 30, 10, "CD+10", CD_plus_10_button  
    OkButton 240, 165, 50, 15
    CancelButton 295, 165, 50, 15
  Text 5, 10, 80, 10, "MNSure Case Number:"
  Text 195, 10, 50, 10, "Verifs due by:"
  Text 5, 25, 300, 10, "If you aren't requesting something, leave that section blank. That way it doesn't case note."
  Text 5, 50, 80, 10, "Household Composition:"
  Text 5, 70, 30, 10, "Income: "
  Text 5, 90, 45, 10, "Residence:"
  Text 5, 110, 55, 10, "Other evidence:"
  Text 215, 145, 70, 10, "Sign your case note:"
EndDialog

DO
	err_msg = ""
	DO
	DIALOG evidence_needed_dialog
		IF ButtonPressed = 0 THEN stopscript
		IF ButtonPressed = CD_plus_10_button THEN verif_due_date = CStr(DateAdd("D", 10, date))
	LOOP UNTIL ButtonPressed = -1
			IF mnsure_case_number = "" THEN err_msg = err_msg & vbCr & "* Please enter the MNSure Case Number."
			IF verif_due_date = "" THEN err_msg = err_msg & vbCr & "* Please enter the document(s) received."
			IF MAXIS_check = 1 AND case_number = "" THEN err_msg = err_msg & vbCr & "* You indicated you wish to case note and TIKL in MAXIS. Please submit a MAXIS case number."
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

objSelection.TypeText "-- RFI SENT --"
objSelection.TypeParagraph()
IF hh_comp <> "" THEN 
	objSelection.TypeText "* Household Composition: " & hh_comp
	objSelection.TypeParagraph()
END IF
IF income <> "" THEN 
	objSelection.TypeText "* Income: " & income
	objSelection.TypeParagraph()
END IF
IF residence <> "" THEN 
	objSelection.TypeText "* Residence: " & residence
	objSelection.TypeParagraph()
END IF
IF other_proofs <> "" THEN 
	objSelection.TypeText "* Other Evidence: " & other_proofs
	objSelection.TypeParagraph()
END IF
IF case_number <> "" THEN 
	objSelection.TypeText "* MAXIS Case Number: " & case_number
	objSelection.TypeParagraph()
	IF case_note_in_maxis_check = 1 THEN 
		objSelection.TypeText "*      Case noted and TIKL'd in MAXIS."
		objSelection.TypeParagraph()
	END IF
END IF
objSelection.TypeText "* Evidence Due By: " & verif_due_date
objSelection.TypeParagraph()
IF sent_arep_checkbox = 1 THEN 
	objSelection.TypeText "* Sent RFI to AREP."
	objSelection.TypeParagraph()
END IF
objSelection.TypeText "***"
objSelection.TypeParagraph()
objSelection.TypeText worker_signature

IF maxis_check = 1 AND case_number <> "" THEN 

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
		CALL write_variable_in_case_note("*** EVIDENCE REQUESTED FOR MNSURE ***")
		CALL write_bullet_and_variable_in_case_note("MNSure Case Number", mnsure_case_number)
		IF hh_comp <> "" THEN CALL write_bullet_and_variable_in_case_note("HH Comp", hh_comp)
		IF income <> "" THEN CALL write_bullet_and_variable_in_case_note("Income", income)
		IF residence <> "" THEN CALL write_bullet_and_variable_in_case_note("Residence", residence)
		IF other_proofs <> "" THEN CALL write_bullet_and_variable_in_case_note("Other Evidence", other_proofs)
		CALL write_variable_in_case_note("---")
		CALL write_bullet_and_variable_in_case_note("Evidence Due Date", verif_due_date)
		CALL write_variable_in_case_note("---")
		CALL write_variable_in_case_note("* TIKL'd 10 days for return.")
		CALL write_variable_in_case_note("---")
		IF sent_arep_checkbox = 1 THEN 
			CALL write_variable_in_case_note("* Sent RFI to AREP.")
			CALL write_variable_in_case_note("---")
		END IF
		CALL write_variable_in_case_note(worker_signature)
	END IF
	
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	CALL create_MAXIS_friendly_date(verif_due_date, 0, 5, 18)
	EMWriteScreen "MNSURE EVIDENCE REQUESTED ON " & date, 9, 3
	EMWriteScreen "SEE MNSURE CASE " & mnsure_case_number & " FOR MORE INFORMATION", 10, 3
	TRANSMIT
	PF3
END IF

'Script end procedure removed by request.
