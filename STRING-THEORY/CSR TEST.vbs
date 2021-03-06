'GATHERING STATS----------------------------------------------------------------------------------------------------
'name_of_script = "NOTES - CSR.vbs"
'start_time = timer

additional_url = "https://raw.githubusercontent.com/RobertFewins-Kalb/Anoka-Specific-Scripts/master/STRING-THEORY/FUNCTIONS.vbs"
SET additional_req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
additional_req.open "GET", additional_url, FALSE							'Attempts to open the FuncLib_URL
additional_req.send													'Sends request
IF additional_req.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute additional_req.responseText								'Executes the script code
ELSE
	MsgBox "The script cannot find the FuncLib"
	stopscript
END IF


'Required for statistical purposes==========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 600          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
next_month = dateadd("m", + 1, date)

footer_month = datepart("m", next_month)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", next_month)
footer_year = "" & footer_year - 2000

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5
Dim row
Dim col

'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS
EMConnect ""
'Searching for the case_number variable
call MAXIS_case_number_finder(case_number)
'Searching for the footer month and footer year
call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Showing the case number dialog
DO
	err_msg = ""
	
BeginDialog case_number_dialog, 0, 0, 166, 265, "Case number dialog"
  EditBox 75, 5, 70, 15, case_number
  EditBox 80, 25, 30, 15, footer_month
  EditBox 115, 25, 30, 15, footer_year
  CheckBox 10, 60, 35, 10, "SNAP", SNAP_checkbox
  CheckBox 95, 60, 30, 10, "HC", HC_checkbox
  CheckBox 10, 80, 100, 10, "Is this an exempt (*) IR?", paperless_checkbox
  EditBox 70, 100, 75, 15, Worker_signature
  ButtonGroup ButtonPressed
    OkButton 35, 135, 50, 15
    CancelButton 95, 135, 50, 15
  Text 10, 10, 50, 10, "Case number:"
  Text 10, 30, 65, 10, "Footer month/year:"
  GroupBox 5, 45, 140, 30, "Programs recertifying"
  Text 10, 105, 60, 10, "Worker Signature"
  GroupBox 10, 165, 155, 75, "Exempt IR checkbox warning:"
  Text 15, 180, 145, 25, "If you select ''Is this an exempt IR'', the case note will only read that the paperless IR was cleared (no case information listed)."
  Text 15, 215, 140, 20, " If you are processing a CSR with SNAP, you should NOT check that option."
EndDialog

	Dialog case_number_dialog
		cancel_confirmation
		If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then err_msg = err_msg & "* You need to type a valid case number."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'Checking for an active MAXIS session
Call check_for_MAXIS(False)

'If "paperless" was checked, the script will put a simple case note in and end.
If paperless_checkbox = 1 then
	call start_a_blank_CASE_NOTE
	Call write_variable_in_case_note("***Cleared paperless IR for " & footer_month & "/" & footer_year & "***")
	Call write_variable_in_case_note("---")
	Call write_variable_in_case_note(worker_signature)
	call script_end_procedure("")
End if

'Navigating to STAT/REVW, checking for error prone cases
call navigate_to_MAXIS_screen("stat", "revw")

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Grabbing SHEL/HEST first, and putting them in this special order that everyone seems to like
call autofill_editbox_from_MAXIS(HH_member_array, "SHEL", SHEL_HEST)
'If SHEL_HEST <> "" then SHEL_HEST = SHEL_HEST & "; "		'this is a temporary fix to resolve issues where a variable is "autofilled" by multiple functions in the same script
call autofill_editbox_from_MAXIS(HH_member_array, "HEST", SHEL_HEST)

'Autofilling HH comp
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)

'Autofilling WREG status
call autofill_editbox_from_MAXIS(HH_member_array, "WREG", notes_on_abawd)

'Autofilling assets
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)

'Autofill DCEX/COEX
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", COEX_DCEX)

'Autofill EI
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)

'Autofill datestamp and UI
call autofill_editbox_from_MAXIS(HH_member_array, "REVW", CSR_datestamp)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

'-----------------Creating text for case note
'Programs recertifying case noting info into variable
If cash_checkbox = 1 then programs_recertifying = programs_recertifying & "cash, "
If HC_checkbox = 1 then programs_recertifying = programs_recertifying & "HC, "
If SNAP_checkbox = 1 then programs_recertifying = programs_recertifying & "SNAP, "
programs_recertifying = trim(programs_recertifying)
if right(programs_recertifying, 1) = "," then programs_recertifying = left(programs_recertifying, len(programs_recertifying) - 1)

'Determining the CSR month for header
CSR_month = footer_month & "/" & footer_year

'Showing the case note dialog
DO
	DO
		Do
			err_msg = ""
			Do
				Do
				
BeginDialog CSR_dialog01, 0, 0, 451, 225, "CSR dialog"
  EditBox 65, 15, 50, 15, CSR_datestamp
  DropListBox 170, 15, 75, 15, "select one..."+chr(9)+"complete"+chr(9)+"incomplete", CSR_status
  EditBox 40, 35, 280, 15, HH_comp
  EditBox 65, 55, 380, 15, earned_income
  EditBox 70, 75, 375, 15, unearned_income
  EditBox 70, 95, 375, 15, notes_on_income
  EditBox 65, 115, 380, 15, notes_on_abawd
  EditBox 40, 135, 405, 15, assets
  EditBox 60, 155, 95, 15, SHEL_HEST
  EditBox 225, 155, 95, 15, COEX_DCEX
  ButtonGroup ButtonPressed
    Pushbutton 340, 205, 50, 15, "Next", next_button
    CancelButton 395, 205, 50, 15
    PushButton 260, 15, 20, 10, "FS", ELIG_FS_button
    PushButton 280, 15, 20, 10, "HC", ELIG_HC_button
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
    PushButton 5, 160, 25, 10, "SHEL/", SHEL_button
    PushButton 30, 160, 25, 10, "HEST:", HEST_button
    PushButton 160, 160, 30, 10, "COEX/", COEX_button
    PushButton 190, 160, 30, 10, "DCEX:", DCEX_button
    PushButton 10, 190, 25, 10, "BUSI", BUSI_button
    PushButton 35, 190, 25, 10, "JOBS", JOBS_button
    PushButton 35, 200, 25, 10, "UNEA", UNEA_button
    PushButton 75, 190, 25, 10, "ACCT", ACCT_button
    PushButton 100, 190, 25, 10, "CARS", CARS_button
    PushButton 125, 190, 25, 10, "CASH", CASH_button
    PushButton 150, 190, 25, 10, "OTHR", OTHR_button
    PushButton 75, 200, 25, 10, "REST", REST_button
    PushButton 100, 200, 25, 10, "SECU", SECU_button
    PushButton 125, 200, 25, 10, "TRAN", TRAN_button
    PushButton 190, 190, 25, 10, "MEMB", MEMB_button
    PushButton 215, 190, 25, 10, "MEMI", MEMI_button
    PushButton 240, 190, 25, 10, "REVW", REVW_button
  GroupBox 255, 5, 50, 25, "ELIG panels:"
  GroupBox 330, 5, 115, 35, "STAT-based navigation:"
  Text 5, 20, 55, 10, "CSR datestamp:"
  Text 125, 20, 40, 10, "CSR status:"
  Text 5, 40, 35, 10, "HH comp:"
  Text 5, 60, 55, 10, "Earned income:"
  Text 5, 80, 60, 10, "Unearned income:"
  Text 5, 100, 60, 10, "Notes on Income:"
  Text 5, 120, 60, 10, "Notes on WREG:"
  Text 5, 140, 30, 10, "Assets:"
  GroupBox 5, 180, 175, 35, "Income and asset panels"
  GroupBox 185, 180, 85, 25, "other STAT panels:"
EndDialog				
				
					Dialog CSR_dialog01
					cancel_confirmation
					IF ButtonPressed = JOBS_button THEN 
						CALL paystubs_received_script()
						'Autofill EI
						call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
						call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
						call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
					END IF
					If ButtonPressed = SIR_mail_button then run "C:\Program Files\Internet Explorer\iexplore.exe https://www.dhssir.cty.dhs.state.mn.us/Pages/Default.aspx"
					'If next_button = pressed THEN msgbox next_button
				Loop until ButtonPressed <> no_cancel_button
				MAXIS_dialog_navigation
			LOOP until ButtonPressed = next_button
			IF CSR_datestamp = "" THEN 														err_msg = err_msg & vbCr & "* Please enter the date the CSR was received."
			IF CSR_status = "select one..." THEN 											err_msg = err_msg & vbCr & "* Please select the status of the CSR."
			IF HH_comp = "" THEN 															err_msg = err_msg & vbCr & "* Please enter household composition information."
			IF earned_income = "" AND unearned_income = "" AND notes_on_income = "" THEN 	err_msg = err_msg & vbCr & "* You must provide some information about income."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		DO
			DO
				DO
				
BeginDialog CSR_dialog02, 0, 0, 451, 260, "CSR dialog"
  EditBox 100, 25, 150, 15, FIAT_reasons
  EditBox 50, 45, 395, 15, other_notes
  EditBox 45, 65, 400, 15, changes
  EditBox 60, 85, 385, 15, verifs_needed
  EditBox 60, 105, 385, 15, actions_taken
  CheckBox 190, 155, 110, 10, "Send forms to AREP?", sent_arep_checkbox
  CheckBox 190, 170, 175, 10, "Check here to case note grant info from ELIG/FS.", grab_FS_info_checkbox
  CheckBox 190, 185, 210, 10, "Check here if CSR and cash supplement were used as a HRF.", HRF_checkbox
  CheckBox 190, 200, 120, 10, "Check here if an eDRS was sent.", eDRS_sent_checkbox
  ButtonGroup ButtonPressed
    PushButton 275, 225, 60, 10, "Previous", previous_button
    OkButton 340, 220, 50, 15
    CancelButton 395, 220, 50, 15
    PushButton 260, 15, 20, 10, "FS", ELIG_FS_button
    PushButton 280, 15, 20, 10, "HC", ELIG_HC_button
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
    PushButton 10, 140, 25, 10, "BUSI", BUSI_button
    PushButton 35, 140, 25, 10, "JOBS", JOBS_button
    PushButton 75, 140, 25, 10, "ACCT", ACCT_button
    PushButton 100, 140, 25, 10, "CARS", CARS_button
    PushButton 125, 140, 25, 10, "CASH", CASH_button
    PushButton 150, 140, 25, 10, "OTHR", OTHR_button
    PushButton 190, 140, 25, 10, "MEMB", MEMB_button
    PushButton 215, 140, 25, 10, "MEMI", MEMI_button
    PushButton 240, 140, 25, 10, "REVW", REVW_button
    PushButton 35, 150, 25, 10, "UNEA", UNEA_button
    PushButton 75, 150, 25, 10, "REST", REST_button
    PushButton 100, 150, 25, 10, "SECU", SECU_button
    PushButton 125, 150, 25, 10, "TRAN", TRAN_button
  EditBox 60, 180, 90, 15, MAEPD_premium
  CheckBox 10, 200, 65, 10, "Emailed MADE?", MADE_checkbox
  ButtonGroup ButtonPressed
    PushButton 80, 200, 65, 10, "SIR mail", SIR_mail_button
  Text 5, 30, 95, 10, "FIAT reasons (if applicable):"
  Text 5, 50, 40, 10, "Other notes:"
  Text 5, 70, 35, 10, "Changes?:"
  Text 5, 90, 50, 10, "Verifs needed:"
  Text 5, 110, 50, 10, "Actions taken:"
  GroupBox 5, 130, 175, 35, "Income and asset panels"
  GroupBox 185, 130, 85, 25, "other STAT panels:"
  GroupBox 5, 170, 150, 45, "If MA-EPD..."
  Text 10, 185, 50, 10, "New premium:"
  GroupBox 255, 5, 50, 25, "ELIG panels:"
  GroupBox 330, 5, 115, 35, "STAT-based navigation:"
EndDialog				
					Dialog CSR_dialog02
					cancel_confirmation
					IF ButtonPressed = JOBS_button THEN paystubs_received_script
					IF ButtonPressed = SIR_mail_button THEN run "C:\Program Files\Internet Explorer\iexplore.exe https://www.dhssir.cty.dhs.state.mn.us/Pages/Default.aspx"
				LOOP UNTIL ButtonPressed <> no_cancel_button
				MAXIS_dialog_navigation
			LOOP UNTIL ButtonPressed = -1 OR ButtonPressed = previous_button
			err_msg = ""
			IF actions_taken = "" THEN 		err_msg = err_msg & vbCr & "* Please indicate the actions you have taken."
			IF err_msg <> "" AND ButtonPressed = -1 THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL err_msg = "" OR ButtonPressed = previous_button
	LOOP WHILE ButtonPressed = previous_button
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

IF grab_FS_info_checkbox = 1 THEN 
	'grabbing information about elig/fs
	call navigate_to_MAXIS_screen("elig", "fs")
	EMReadScreen FSPR_check, 4, 3, 48
	If FSPR_check <> "FSPR" then
		MsgBox "The script couldn't find ELIG/FS. It will now jump to case note."
	Else
		EMWriteScreen "FSSM", 19, 70
		transmit
		EMReadScreen FSSM_line_01, 37, 13, 44
		EMReadScreen FSSM_line_02, 37, 8, 3
		EMReadScreen FSSM_line_03, 37, 10, 3
	End if
END IF

'Writing the case note to MAXIS----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
call write_variable_in_case_note("***" & CSR_month & " CSR received " & CSR_datestamp & ": " & CSR_status & "***")
call write_bullet_and_variable_in_case_note("Programs recertifying", programs_recertifying)
call write_bullet_and_variable_in_case_note("HH comp", HH_comp)
call write_bullet_and_variable_in_case_note("Earned income", earned_income)
call write_bullet_and_variable_in_case_note("Unearned income", unearned_income)
call write_bullet_and_variable_in_case_note("Notes on Income", notes_on_income)
call write_bullet_and_variable_in_case_note("ABAWD Notes", notes_on_abawd)
call write_bullet_and_variable_in_case_note("Assets", assets)
call write_bullet_and_variable_in_case_note("SHEL/HEST", SHEL_HEST)
call write_bullet_and_variable_in_case_note("COEX/DCEX", COEX_DCEX)
call write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
call write_bullet_and_variable_in_case_note("Other notes", other_notes)
call write_bullet_and_variable_in_case_note("Changes", changes)
If HRF_checkbox = checked then call write_variable_in_case_note("* CSR and cash supplement used as HRF.")
If eDRS_sent_checkbox = checked then call write_variable_in_case_note("* eDRS sent.")
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
call write_bullet_and_variable_in_case_note("MA-EPD premium", MAEPD_premium)
If MADE_checkbox = checked then call write_variable_in_case_note("* Emailed MADE through DHS-SIR.")
call write_variable_in_case_note("---")
If grab_FS_info_checkbox = 1 AND FSPR_check = "FSPR" then
	call write_variable_in_case_note("   " & FSSM_line_01)
	call write_variable_in_case_note("   " & FSSM_line_02)
	call write_variable_in_case_note("   " & FSSM_line_03)
	call write_variable_in_case_note("---")
End if
call write_variable_in_case_note(worker_signature)

call script_end_procedure("")
