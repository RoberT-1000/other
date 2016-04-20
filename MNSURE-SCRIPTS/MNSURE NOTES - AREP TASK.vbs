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

'=====CUSTOM FUNCTION: NAVIGATING TO STAT/AREP AND PULL AREP INFO=====
FUNCTION grab_MAXIS_AREP_info(case_number, arep_name, addrLine1, addrLine2, addrCity, addrState, addrZip, addrPhone)
	CALL check_for_MAXIS(False)
	case_number = case_number
	CALL navigate_to_MAXIS_screen("STAT", "AREP")
	EMReadScreen arep_name, 37, 4, 32
	
	arep_name = replace(arep_name, "_", "")
	IF arep_name = "" THEN 
		MsgBox "STAT/AREP is blank on MAXIS case " & case_number & ". Please change your case number or enter AREP information manually."
	ELSE
		EMReadScreen addrLine1, 22, 5, 32
		addrLine1 = replace(addrLine1, "_", "")
		EMReadScreen addrLine2, 22, 6, 32
		addrLine2 = replace(addrLine2, "_", "")
		EMReadScreen addrCity, 15, 7, 32
		addrCity = replace(addrCity, "_", "")
		EMReadScreen addrState, 2, 7, 55
		EMReadScreen addrZip, 5, 7, 64
		EMReadScreen addrPhone, 14, 8, 34
		addrPhone = replace(addrPhone, " ) ", "-")
		addrPhone = replace(addrPhone, " ", "-")
		IF addrPhone = "___-___-____" THEN addrPhone = ""
		
	END IF
END FUNCTION

'===== THE DIALOG =====
BeginDialog mnsure_areg_dlg, 0, 0, 231, 325, "MNSure AREP"
  EditBox 95, 25, 55, 15, look_up_maxis_case_number
  ButtonGroup ButtonPressed
    PushButton 160, 25, 30, 15, "GO!", nav_to_arep_button
  EditBox 85, 80, 120, 15, arep_name
  EditBox 85, 100, 120, 15, arep_addr_line1
  EditBox 85, 120, 120, 15, arep_addr_line2
  EditBox 85, 140, 60, 15, arep_addr_city
  EditBox 150, 140, 20, 15, arep_addr_state
  EditBox 175, 140, 30, 15, arep_addr_zipCode
  EditBox 85, 160, 75, 15, arep_phone
  EditBox 95, 190, 80, 15, mnsure_case_number
  EditBox 60, 210, 160, 15, other_notes
  CheckBox 10, 235, 160, 10, "Check HERE to case note and TIKL in MAXIS.", maxis_case_note_check
  EditBox 90, 250, 65, 15, case_number
  EditBox 90, 280, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 125, 305, 50, 15
    CancelButton 175, 305, 50, 15
  Text 20, 30, 70, 10, "MAXIS Case Number"
  Text 20, 85, 45, 10, "AREP Name:"
  Text 20, 105, 55, 10, "AREP Address:"
  Text 20, 165, 50, 10, "AREP Phone:"
  Text 10, 195, 85, 10, "Integrated Case Number:"
  Text 10, 215, 45, 10, "Other Notes:"
  Text 35, 255, 50, 10, "Case Number:"
  Text 15, 285, 70, 10, "Sign your case note: "
  GroupBox 10, 10, 210, 45, "Pulling AREP Info from MAXIS"
  GroupBox 10, 65, 210, 115, "AREP Info"
EndDialog


'=====THE SCRIPT=====
EMConnect ""
DO
	err_msg = "" 
	DIALOG mnsure_areg_dlg
		IF ButtonPressed = 0 THEN stopscript
		IF ButtonPressed = nav_to_arep_button AND look_up_maxis_case_number <> "" THEN 
			case_number = look_up_maxis_case_number
			CALL grab_MAXIS_AREP_info(case_number, arep_name, arep_addr_line1, arep_addr_line2, arep_addr_city, arep_addr_state, arep_addr_zipCode, arep_phone)
		END IF
		IF arep_name = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please provide an AREP name."
		IF maxis_case_note_check = 1 AND case_number = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter a valid MAXIS case number for the script to interact with MAXIS."
		IF worker_signature = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = "" AND ButtonPressed = -1
		
'Now it creates a Word document to store all of the active claims.
Set objWord = CreateObject("Word.Application")
objWord.Visible = true
set objDoc = objWord.Documents.add()
Set objSelection = objWord.Selection
objSelection.ParagraphFormat.LineSpacing = 12
objSelection.Range.ParagraphFormat.SpaceAfter = 0

objSelection.TypeText "-- AREP ON FILE --"
objSelection.TypeParagraph()
objSelection.TypeText "* AREP Name: " & arep_name
objSelection.TypeParagraph()
IF arep_addr_line1 <> "" THEN 
	objSelection.TypeText "* Address..."
	objSelection.TypeParagraph
	objSelection.TypeText "*          " & arep_addr_line1
	objSelection.TypeParagraph()
	IF arep_addr_line2 <> "" THEN
		objSelection.TypeText "*          " & arep_addr_line2
		objSelection.TypeParagraph()
	END IF
	objSelection.TypeText "*          " & arep_addr_city & ", " & arep_addr_state & " " & arep_addr_zipCode
	objSelection.TypeParagraph()
END IF
IF phone_number <> "" THEN 
	objSelection.TypeText "* Phone Number: " & phone_number
	objSelection.TypeParagraph()
END IF
IF other_notes <> "" THEN 
	objSelection.TypeText "* Other Notes: " & other_notes
	objSelection.TypeParagraph()
END IF
objSelection.TypeText "***"
objSelection.TypeParagraph()
objSelection.TypeText worker_signature


'=====Interacting with MAXIS=====
IF maxis_case_note_check = 1 THEN 
	CALL check_for_MAXIS(False)
	start_a_blank_CASE_NOTE
	
	CALL write_variable_in_CASE_NOTE("*** UPDATED AREP IN MNSURE CASE " & mnsure_case_number & " ***")
	CALL write_bullet_and_variable_in_CASE_NOTE("AREP Name", arep_name)
	IF arep_addr_line1 <> "" THEN
		CALL write_variable_in_CASE_NOTE("* AREP Address:  " & arep_addr_line1)
		IF arep_addr_line2 <> "" THEN CALL write_variable_in_CASE_NOTE("                 " & arep_addr_line2)
		CALL write_variable_in_CASE_NOTE("                 " & arep_addr_city & ", " & arep_addr_state & " " & arep_addr_zipCode)
	END IF
	IF arep_phone <> "" THEN CALL write_bullet_and_variable_in_CASE_NOTE("AREP Phone", arep_phone)
	IF other_notes <> "" THEN CALL write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
	CALL write_variable_in_CASE_NOTE("---")
	CALL write_variable_in_CASE_NOTE(worker_signature)
END IF

	
	
	
	
	
	
	
	
	
	
	
