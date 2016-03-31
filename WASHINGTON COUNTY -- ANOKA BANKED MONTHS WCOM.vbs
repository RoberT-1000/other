'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "MEMOS - BANKED MONTHS WCOM.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'Dialogs
BeginDialog case_number_dlg, 0, 0, 211, 195, "Case Number Dialog"
  EditBox 70, 10, 60, 15, case_number
  EditBox 70, 50, 30, 15, approval_month
  EditBox 160, 50, 30, 15, approval_year
  EditBox 120, 115, 60, 15, appointment_date
  ButtonGroup ButtonPressed
    OkButton 95, 175, 50, 15
    CancelButton 150, 175, 50, 15
  Text 10, 15, 55, 10, "Case Number: "
  Text 10, 55, 55, 10, "Approval Month:"
  Text 105, 55, 50, 10, "Approval Year:"
  Text 10, 70, 185, 10, "* Use these fields when adding WCOM to notice."
  Text 15, 120, 100, 10, "WF1 Appt Date (MM/DD/YY):"
  Text 10, 135, 180, 15, "* Use this field when setting a manual WF1 referral."
  GroupBox 5, 35, 195, 55, "WCOM Approval Month"
  GroupBox 5, 100, 195, 55, "WF1 Appointment Date"
EndDialog

BeginDialog banked_months_menu_dialog, 0, 0, 356, 170, "Banked Months Main Menu"
  ButtonGroup ButtonPressed
    PushButton 10, 15, 65, 10, "WF1 Referral", wf1_button
    PushButton 10, 55, 90, 10, "All Banked Months Used", banked_months_used_button
    PushButton 10, 80, 90, 10, "Banked Months Notifier", banked_months_notifier
    PushButton 10, 105, 90, 10, "Closing for E/T Non-Coop", e_t_non_coop_button
    CancelButton 300, 150, 50, 15
  Text 80, 15, 260, 10, "-- Use this script to generate a WF1 referral."
  Text 110, 55, 230, 20, "-- Use this script when a client's SNAP is closing because they used all their eligible banked months."
  Text 110, 80, 225, 20, "-- Use this script to add a WCOM to a notice notifying the client they may be eligible for banked months."
  Text 110, 105, 225, 20, "-- Use this script to add a WCOM to a client's closing notice to inform them they are closing on banked months for SNAP E&T Non-Coop."
  GroupBox 5, 40, 345, 90, "WCOM"
EndDialog

BeginDialog wcom_only_banked_months_menu_dialog, 0, 0, 356, 140, "Banked Months WCOMs"
  ButtonGroup ButtonPressed
    PushButton 10, 25, 90, 10, "All Banked Months Used", banked_months_used_button
    PushButton 10, 50, 90, 10, "Banked Months Notifier", banked_months_notifier
    PushButton 10, 75, 90, 10, "Closing for E/T Non-Coop", e_t_non_coop_button
    CancelButton 300, 120, 50, 15
  Text 110, 25, 230, 20, "-- Use this script when a client's SNAP is closing because they used all their eligible banked months."
  Text 110, 50, 230, 20, "-- Use this script to add a WCOM to a notice notifying the client they may be eligible for banked months."
  Text 110, 75, 235, 20, "-- Use this script to add a WCOM to a client's closing notice to inform them they are closing on banked months for Employment Services Non-Coop."
  GroupBox 5, 10, 345, 90, "WCOM"
EndDialog


BeginDialog super_user_dialog, 0, 0, 251, 70, "Banked Months Main Menu"
  DropListBox 105, 15, 135, 15, "Select one:"+chr(9)+"TIKL RE: FSET Non-Coop"+chr(9)+"TIKL RE: CL is active with E&T"+chr(9)+"TIKL RE: CL not active with E&T"+chr(9)+"Generate WCOM"+chr(9)+"WF1 Referral", run_mode
  ButtonGroup ButtonPressed
    OkButton 5, 50, 50, 15
    CancelButton 55, 50, 50, 15
  Text 10, 15, 85, 10, "Please select run mode:"
EndDialog


'--- The script -----------------------------------------------------------------------------------------------------------------

EMConnect ""

'Grabbing user ID to validate user of script. Only some users are allowed to use this script.
Set objNet = CreateObject("WScript.NetWork") 
user_ID_for_validation = ucase(objNet.UserName)

'For the super users
IF user_ID_for_validation = "RAKALB" OR _
	user_ID_for_validation = "SLCARDA" OR _ 
	user_ID_for_validation = "EABUELOW" OR _ 
	user_ID_for_validation = "EAKUNZMA" OR _
	user_ID_for_validation = "CDPOTTER"	THEN 
	DO
		DIALOG super_user_dialog
			IF ButtonPressed = 0 THEN stopscript
			IF run_mode = "Select one:" THEN MsgBox "Try again. Please pick a run mode."
	LOOP UNTIL run_mode <> "Select one:"
	
	IF run_mode = "TIKL RE: FSET Non-Coop" THEN 
		CALL run_another_script("Q:\Blue Zone Scripts\Public assistance script files\Script Files\County Customized\BULK - FSET non-compliance TIKLer.vbs")
		stopscript
		
	ELSEIF run_mode = "TIKL RE: CL is active with E&T" THEN
		coop_mode = true
		noncoop_mode = false
		CALL run_another_script("Q:\Blue Zone Scripts\Public assistance script files\Script Files\County Customized\BULK - BANKED MONTHS TIKLER.vbs")
		stopscript
		
	ELSEIF run_mode = "TIKL RE: CL not active with E&T" THEN 
		coop_mode = false
		noncoop_mode = true
		CALL run_another_script("Q:\Blue Zone Scripts\Public assistance script files\Script Files\County Customized\BULK - BANKED MONTHS TIKLER.vbs")
		stopscript
		
	ELSEIF run_mode = "Generate WCOM" THEN 
		DO
			err_msg = ""
			dialog case_number_dlg
			cancel_confirmation
			IF case_number = "" THEN err_msg = "* Please enter a case number" & vbNewLine
			IF len(approval_month) <> 2 THEN err_msg = err_msg & "* Please enter your month in MM format." & vbNewLine
			IF len(approval_year) <> 2 THEN err_msg = err_msg & "* Please enter your year in YY format." & vbNewLine
			IF err_msg <> "" THEN msgbox err_msg
		LOOP until err_msg = ""	
		
		DIALOG wcom_only_banked_months_menu_dialog
			IF ButtonPressed = 0 THEN stopscript
			
			'This is the WCOM for when the client has used all their banked months.
			IF ButtonPressed = banked_months_used_button THEN
				call navigate_to_MAXIS_screen("spec", "wcom")
				
				EMWriteScreen approval_month, 3, 46
				EMWriteScreen approval_year, 3, 51
				transmit
				
				DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
					EMReadScreen more_pages, 8, 18, 72
					IF more_pages = "MORE:  -" THEN PF7
				LOOP until more_pages <> "MORE:  -"
				
				read_row = 7
				DO
					waiting_check = ""
					EMReadscreen prog_type, 2, read_row, 26
					EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
					If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
						EMSetcursor read_row, 13
						EMSendKey "x"
						Transmit
						pf9
						EMSetCursor 03, 15
						CALL write_variable_in_SPEC_MEMO("You have been receiving SNAP banked months. Your SNAP is closing for using all available banked months. If you meet one of the exemptions listed above AND all other eligibility factors you may still be eligible for SNAP. Please contact your financial worker if you have questions.")
						PF4
						PF3
						WCOM_count = WCOM_count + 1
						exit do
					ELSE
						read_row = read_row + 1
					END IF
					IF read_row = 18 THEN
						PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
						read_row = 7
					End if
				LOOP until prog_type = "  "
				
				wcom_type = "all banked months"
				
			'This is the WCOM for when the client is closing for ABAWD and is being notified that they could be eligible for banked months.
			ELSEIF ButtonPressed = banked_months_notifier THEN 
				call navigate_to_MAXIS_screen("spec", "wcom")
				
				EMWriteScreen approval_month, 3, 46
				EMWriteScreen approval_year, 3, 51
				transmit
				
				DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
					EMReadScreen more_pages, 8, 18, 72
					IF more_pages = "MORE:  -" THEN PF7
				LOOP until more_pages <> "MORE:  -"
				
				read_row = 7
				DO
					waiting_check = ""
					EMReadscreen prog_type, 2, read_row, 26
					EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
					If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
						EMSetcursor read_row, 13
						EMSendKey "x"
						Transmit
						pf9
						EMSetCursor 03, 15
						CALL write_variable_in_SPEC_MEMO("You have used all of your available ABAWD months. You may be eligible for SNAP banked months if you are cooperating with Employment Services. Please contact your financial worker if you have questions.")
						PF4
						PF3
						WCOM_count = WCOM_count + 1
						exit do
					ELSE
						read_row = read_row + 1
					END IF
					IF read_row = 18 THEN
						PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
						read_row = 7
					End if
				LOOP until prog_type = "  "
				
				wcom_type = "banked months notifier"
		
			'This is the WCOM for when the client is closing on banked months for E&T Non-Coop
			ELSEIF ButtonPressed = e_t_non_coop_button THEN 
			
				DO
					hh_member = InputBox("Please enter the name of the client that is closing for E&T Non-Coop...")
					confirmation_msg = MsgBox("Please confirm to add the client's name to the WCOM: " & vbCr & vbCr & hh_member & " is closing on banked months for SNAP E&T Non-Cooperation." & vbCr & vbCr & "Is this correct? Press YES to continue. Press NO to re-enter the client's name. Press CANCEL to stop the script.", vbYesNoCancel)
					IF confirmation_msg = vbCancel THEN stopscript
				LOOP UNTIL confirmation_msg = vbYes
				
				call navigate_to_MAXIS_screen("spec", "wcom")
				
				EMWriteScreen approval_month, 3, 46
				EMWriteScreen approval_year, 3, 51
				transmit
				
				DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
					EMReadScreen more_pages, 8, 18, 72
					IF more_pages = "MORE:  -" THEN PF7
				LOOP until more_pages <> "MORE:  -"
				
				read_row = 7
				DO
					waiting_check = ""
					EMReadscreen prog_type, 2, read_row, 26
					EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
					If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
						EMSetcursor read_row, 13
						EMSendKey "x"
						Transmit
						pf9
						EMSetCursor 03, 15
						CALL write_variable_in_SPEC_MEMO("You have been receiving SNAP banked months. Your SNAP case is closing because " & hh_member & " did not meet the requirements of working with Employment and Training. If you feel you have Good Cause for not cooperating with this requirement please contact your financial worker before your SNAP closes. If your SNAP closes for not cooperating with Employment and Training you will not be eligible for future banked months. If you meet an exemption listed above AND all other eligibility factors you may be eligible for SNAP. If you have questions please contact your financial worker.")
						PF4
						PF3
						WCOM_count = WCOM_count + 1
						exit do
					ELSE
						read_row = read_row + 1
					END IF
					IF read_row = 18 THEN
						PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
						read_row = 7
					End if
				LOOP until prog_type = "  "
				
				wcom_type = "non coop"
				
			END IF
		
		'Outcome ---------------------------------------------------------------------------------------------------------------------
		
		If WCOM_count = 0 THEN  'if no waiting FS notice is found
			script_end_procedure("No Waiting FS elig results were found in this month for this HH member.")
		ELSE 					'If a waiting FS notice is found
			'Case note
			start_a_blank_case_note
			call write_variable_in_CASE_NOTE("---WCOM added regarding banked months---")
			IF wcom_type = "all banked months" THEN 
				CALL write_variable_in_CASE_NOTE("* WCOM added because client all eligible banked months have been used.")
			ELSEIF wcom_type = "non coop" THEN
				CALL write_variable_in_CASE_NOTE("* Banked months ending for SNAP E & T non-coop.")
			ELSEIF wcom_type = "banked months notifier" THEN 
				CALL write_variable_in_CASE_NOTE("* Client has used ABAWD counted months and MAY be eligible for banked months. Eligibility questions should be directed to financial worker.")
			END IF
			
			call write_variable_in_CASE_NOTE("---")
			IF worker_signature <> "" THEN 
				call write_variable_in_CASE_NOTE(worker_signature)
			ELSE
				worker_signature = InputBox("Please sign your case note...")
				CALL write_variable_in_CASE_NOTE(worker_signature)
			END IF
		END IF
		
		script_end_procedure("")
		
	ELSEIF run_mode = "WF1 Referral" THEN 
		DO
			err_msg = ""
			dialog case_number_dlg
			cancel_confirmation
			IF case_number = "" THEN err_msg = "* Please enter a case number" & vbNewLine
			IF appointment_date = "" OR IsDate(appointment_date) = FALSE THEN err_msg = err_msg & "* Please enter a valid appointment date." & vbNewLine
			IF err_msg <> "" THEN msgbox "*** NOTICE!!! ***" & vbCr & vbCr & err_msg & "Please resolve for the script to continue."
		LOOP until err_msg = ""	
		
		CALL check_for_MAXIS(false)
		
		MsgBox "*** NOTICE!!! ***" & vbCr & vbCr & "You will be asked to select the household members to refer. Please select all members that are being assigned the appointment date " & appointment_date & "."
		CALL HH_member_custom_dialog(HH_member_array)
	
		CALL check_for_MAXIS(false)

		Call navigate_to_MAXIS_screen("INFC", "WF1M")			'navigates to WF1M to create the manual referral'
		EMWriteScreen "01", 4, 47								'using "OTHER" code for WF1 referral
		
		WF1M_row = 8
		FOR EACH hh_memb IN HH_member_array
			EMWriteScreen hh_memb, WF1M_row, 9
			EMWriteScreen "FS", WF1M_row, 46
			EMWriteScreen "X", WF1M_row, 53
			CALL create_MAXIS_friendly_date(appointment_date, 0, WF1M_row, 65)
			WF1M_row = WF1M_row + 1
		NEXT
		
		'transmitting to send WF1M to select center
		transmit
		
		FOR EACH hh_memb IN HH_member_array
			EMSendKey "X"
			transmit
		NEXT
		
		'Now we are back at the WF1M screen
		EMWriteScreen "WF1 REFERRAL FOR ABAWD BANKED MONTHS; ", 17, 6
		EMSetCursor 17, 44
	
		closing_notice = "************************************ NOTICE!!! ************************************" & vbCr & vbCr & "* Please enter additional information before submitting the referral." & vbCr & _ 
							"* Enter the months being used as banked months." & vbCr & _ 
							"* Enter any personal exemptions the client may have which could grant them eligibility for additional banked months, including: " & vbCr & _
							vbTab & "-- Veteran of Armed Forces" & vbCr & _
							vbTab & "-- Homelessness" & vbCr & _
							vbTab & "-- Aging out of Foster Care" & vbCr & _
							vbTab & "-- Child Aging out of MFIP, and" & vbCr & _
							vbTab & "-- Victim of Domestic Violence" & vbCr & vbCr & _
							"* Press PF3 to submit the referral." & vbCr & vbCr & _
							"* Please be sure to update STAT/MISC with banked month information. We are not able to effectively track banked months without that panel being updated." & vbCr & vbCr & _
							"* Send the SNAP Orientation Letter (GEN 281) in Compass Pilot." & vbCr & vbCr & _
							"***********************************************************************************"

		script_end_procedure(closing_notice)
	END IF
END IF

'>>>>> THIS IS THE GENERIC RUN MODE FOR ALL USERS FOR WCOMs <<<<<
call MAXIS_case_number_finder(case_number)
approval_month = DatePart("M", (DateAdd("M", 1, date)))
IF len(approval_month) = 1 THEN 
	approval_month = "0" & approval_month
ELSE
	approval_month = Cstr(approval_month)
END IF
approval_year = Right(DatePart("YYYY", (DateAdd("M", 1, date))), 2)

CALL check_for_MAXIS(false)

DIALOG banked_months_menu_dialog
	IF ButtonPressed = 0 THEN stopscript
	
	'This is the WCOM for when the client has used all their banked months.
	IF ButtonPressed = banked_months_used_button THEN
		call navigate_to_MAXIS_screen("spec", "wcom")
		
		EMWriteScreen approval_month, 3, 46
		EMWriteScreen approval_year, 3, 51
		transmit
		
		DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
			EMReadScreen more_pages, 8, 18, 72
			IF more_pages = "MORE:  -" THEN PF7
		LOOP until more_pages <> "MORE:  -"
		
		read_row = 7
		DO
			waiting_check = ""
			EMReadscreen prog_type, 2, read_row, 26
			EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
			If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
				EMSetcursor read_row, 13
				EMSendKey "x"
				Transmit
				pf9
				EMSetCursor 03, 15
				CALL write_variable_in_SPEC_MEMO("You have been receiving SNAP banked months. Your SNAP is closing for using all available banked months. If you meet one of the exemptions listed above AND all other eligibility factors you may still be eligible for SNAP. Please contact your financial worker if you have questions.")
				PF4
				PF3
				WCOM_count = WCOM_count + 1
				exit do
			ELSE
				read_row = read_row + 1
			END IF
			IF read_row = 18 THEN
				PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
				read_row = 7
			End if
		LOOP until prog_type = "  "
		
		wcom_type = "all banked months"
		
	'This is the WCOM for when the client is closing for ABAWD and is being notified that they could be eligible for banked months.
	ELSEIF ButtonPressed = banked_months_notifier THEN 
		call navigate_to_MAXIS_screen("spec", "wcom")
		
		EMWriteScreen approval_month, 3, 46
		EMWriteScreen approval_year, 3, 51
		transmit
		
		DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
			EMReadScreen more_pages, 8, 18, 72
			IF more_pages = "MORE:  -" THEN PF7
		LOOP until more_pages <> "MORE:  -"
		
		read_row = 7
		DO
			waiting_check = ""
			EMReadscreen prog_type, 2, read_row, 26
			EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
			If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
				EMSetcursor read_row, 13
				EMSendKey "x"
				Transmit
				pf9
				EMSetCursor 03, 15
				CALL write_variable_in_SPEC_MEMO("You have used all of your available ABAWD months. You may be eligible for SNAP banked months if you are cooperating with Employment Services. Please contact your financial worker if you have questions.")
				PF4
				PF3
				WCOM_count = WCOM_count + 1
				exit do
			ELSE
				read_row = read_row + 1
			END IF
			IF read_row = 18 THEN
				PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
				read_row = 7
			End if
		LOOP until prog_type = "  "
		
		wcom_type = "banked months notifier"

	'This is the WCOM for when the client is closing on banked months for E&T Non-Coop
	ELSEIF ButtonPressed = e_t_non_coop_button THEN 
	
		DO
			hh_member = InputBox("Please enter the name of the client that is closing for E&T Non-Coop...")
			confirmation_msg = MsgBox("Please confirm to add the client's name to the WCOM: " & vbCr & vbCr & hh_member & " is closing on banked months for SNAP E&T Non-Cooperation." & vbCr & vbCr & "Is this correct? Press YES to continue. Press NO to re-enter the client's name. Press CANCEL to stop the script.", vbYesNoCancel)
			IF confirmation_msg = vbCancel THEN stopscript
		LOOP UNTIL confirmation_msg = vbYes
		
		call navigate_to_MAXIS_screen("spec", "wcom")
		
		EMWriteScreen approval_month, 3, 46
		EMWriteScreen approval_year, 3, 51
		transmit
		
		DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
			EMReadScreen more_pages, 8, 18, 72
			IF more_pages = "MORE:  -" THEN PF7
		LOOP until more_pages <> "MORE:  -"
		
		read_row = 7
		DO
			waiting_check = ""
			EMReadscreen prog_type, 2, read_row, 26
			EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
			If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
				EMSetcursor read_row, 13
				EMSendKey "x"
				Transmit
				pf9
				EMSetCursor 03, 15
				CALL write_variable_in_SPEC_MEMO("You have been receiving SNAP banked months. Your SNAP case is closing because " & hh_member & " did not meet the requirements of working with Employment and Training. If you feel you have Good Cause for not cooperating with this requirement please contact your financial worker before your SNAP closes. If your SNAP closes for not cooperating with Employment and Training you will not be eligible for future banked months. If you meet an exemption listed above AND all other eligibility factors you may be eligible for SNAP. If you have questions please contact your financial worker.")
				PF4
				PF3
				WCOM_count = WCOM_count + 1
				exit do
			ELSE
				read_row = read_row + 1
			END IF
			IF read_row = 18 THEN
				PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
				read_row = 7
			End if
		LOOP until prog_type = "  "
		
		wcom_type = "non coop"
	
	'This is the run mode for sending a WF1 referral.
	ELSEIF ButtonPressed = wf1_button THEN 
		DO
			err_msg = ""
			dialog case_number_dlg
			cancel_confirmation
			IF case_number = "" THEN err_msg = "* Please enter a case number" & vbNewLine
			IF appointment_date = "" OR IsDate(appointment_date) = FALSE THEN err_msg = err_msg & "* Please enter a valid appointment date." & vbNewLine
			IF err_msg <> "" THEN msgbox "*** NOTICE!!! ***" & vbCr & vbCr & err_msg & "Please resolve for the script to continue."
		LOOP until err_msg = ""	
		
		CALL check_for_MAXIS(false)
		
		MsgBox "*** NOTICE!!! ***" & vbCr & vbCr & "You will be asked to select the household members to refer. Please select all members that are being assigned the appointment date " & appointment_date & "."
		CALL HH_member_custom_dialog(HH_member_array)
	
		CALL check_for_MAXIS(false)

		Call navigate_to_MAXIS_screen("INFC", "WF1M")			'navigates to WF1M to create the manual referral'
		EMWriteScreen "01", 4, 47								'using "OTHER" code for WF1 referral
		
		WF1M_row = 8
		FOR EACH hh_memb IN HH_member_array
			EMWriteScreen hh_memb, WF1M_row, 9
			EMWriteScreen "FS", WF1M_row, 46
			EMWriteScreen "X", WF1M_row, 53
			CALL create_MAXIS_friendly_date(appointment_date, 0, WF1M_row, 65)
			WF1M_row = WF1M_row + 1
		NEXT
		
		'transmitting to send WF1M to select center
		transmit
		
		FOR EACH hh_memb IN HH_member_array
			EMSendKey "X"
			transmit
		NEXT
		
		'Now we are back at the WF1M screen
		EMWriteScreen "WF1 REFERRAL FOR ABAWD BANKED MONTHS; ", 17, 6
		EMSetCursor 17, 44
	
		closing_notice = "************************************ NOTICE!!! ************************************" & vbCr & vbCr & "* Please enter additional information before submitting the referral." & vbCr & _ 
							"* Enter the months being used as banked months." & vbCr & _ 
							"* Enter any personal exemptions the client may have which could grant them eligibility for additional banked months, including: " & vbCr & _
							vbTab & "-- Veteran of Armed Forces" & vbCr & _
							vbTab & "-- Homelessness" & vbCr & _
							vbTab & "-- Aging out of Foster Care" & vbCr & _
							vbTab & "-- Child Aging out of MFIP, and" & vbCr & _
							vbTab & "-- Victim of Domestic Violence" & vbCr & vbCr & _
							"* Press PF3 to submit the referral." & vbCr & vbCr & _
							"* Please be sure to update STAT/MISC with banked month information. We are not able to effectively track banked months without that panel being updated." & vbCr & vbCr & _
							"* Send the SNAP Orientation Letter (GEN 281) in Compass Pilot." & vbCr & vbCr & _
							"***********************************************************************************"

		script_end_procedure(closing_notice)
	
	END IF

'Outcome ---------------------------------------------------------------------------------------------------------------------

If WCOM_count = 0 THEN  'if no waiting FS notice is found
	script_end_procedure("No Waiting FS elig results were found in this month for this HH member.")
ELSE 					'If a waiting FS notice is found
	'Case note
	start_a_blank_case_note
	call write_variable_in_CASE_NOTE("---WCOM added regarding banked months---")
	IF wcom_type = "all banked months" THEN 
		CALL write_variable_in_CASE_NOTE("* WCOM added because client all eligible banked months have been used.")
	ELSEIF wcom_type = "non coop" THEN
		CALL write_variable_in_CASE_NOTE("* Banked months ending for SNAP E & T non-coop.")
	ELSEIF wcom_type = "banked months notifier" THEN 
		CALL write_variable_in_CASE_NOTE("* Client has used ABAWD counted months and MAY be eligible for banked months. Eligibility questions should be directed to financial worker.")
	END IF
	
	call write_variable_in_CASE_NOTE("---")
	IF worker_signature <> "" THEN 
		call write_variable_in_CASE_NOTE(worker_signature)
	ELSE
		worker_signature = InputBox("Please sign your case note...")
		CALL write_variable_in_CASE_NOTE(worker_signature)
	END IF
END IF

script_end_procedure("")
