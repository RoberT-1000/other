'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - pull cases into Excel-revised"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"

SET req = CreateObject("Msxml2.XMLHttp.6.0") 'Creates an object to get a URL
req.open "GET", url, FALSE	'Attempts to open the URL
req.send 'Sends request

IF req.Status = 200 THEN	'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject") 'Creates an FSO
	Execute req.responseText 'Executes the script code
ELSE	'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox "Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
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
	"URL: " & url
	script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

'----------FUNCTIONS----------
FUNCTION check_panels_function(x, panel_status)
	FOR EACH hh_person IN x
		CALL navigate_to_screen("STAT", "HEST")
		errr_screen_check
			IF hh_person <> "01" THEN
				EMWriteScreen hh_person, 20, 76
				transmit
			END IF
			EMReadScreen hest_info, 1, 2, 73
			IF hest_info <> "0" THEN panel_status = panel_status & "HEST"
			IF hest_info = "0" THEN 
				EMWriteScreen "SHEL", 20, 71
				transmit
				IF hh_person <> "01" THEN
					EMWriteScreen hh_person, 20, 76
					transmit
				END IF
				EMReadScreen shel_info, 1, 2, 73
				IF shel_info <> "0" THEN panel_status = "SHEL"
				IF shel_info = "0" THEN
					EMWriteScreen "COEX", 20, 71
					transmit
					IF hh_person <> "01" THEN
						EMWriteScreen hh_person, 20, 76
						transmit
					END IF
					EMReadScreen coex_info, 1, 2, 73
					IF coex_info <> "0" THEN panel_status = "COEX"
					IF coex_info = "0" THEN
						EMWriteScreen "DCEX", 20, 71
						transmit
						IF hh_person <> "01" THEN
							EMWriteScreen hh_person, 20, 76
							transmit
						END IF
						EMReadScreen dcex_info, 1, 2, 73
						IF dcex_info = "0" THEN
							EMWriteScreen "BUSI", 20, 71
							transmit
							IF hh_person <> "01" THEN
								EMWriteScreen hh_person, 20, 76
								transmit
							END IF
							EMReadScreen busi_info, 1, 2, 73
							EMReadScreen busi_end_date, 8, 5, 71
							IF busi_info = "0" OR busi_end_date <> "__ __ __" THEN 
								EMWriteScreen "UNEA", 20, 71
								transmit
								IF hh_person <> "01" THEN
									EMWriteScreen hh_person, 20, 76
									transmit
								END IF
								EMReadScreen unea_info, 1, 2, 73
								IF unea_info <> "0" THEN
									DO
										EMReadScreen unea_end_date, 8, 9, 68
										IF unea_end_date <> "__ __ __" THEN
											transmit
											EMReadScreen valid_command, 21, 24, 2
										END IF
									LOOP UNTIL valid_command = "ENTER A VALID COMMAND" OR unea_end_date = "__ __ __"
								END IF
								IF (unea_info <> "0" AND unea_end_date = "__ __ __") THEN panel_status = "UNEA"
								IF (unea_info <> "0" AND valid_command = "ENTER A VALID COMMAND") OR unea_info = "0" THEN
									CALL navigate_to_screen("STAT", "JOBS")
									IF hh_person <> "01" THEN
										EMWriteScreen hh_person, 20, 76
										transmit
									END IF
									EMReadScreen jobs_info, 1, 2, 73
									IF jobs_info <> "0" THEN
										DO
											EMReadScreen jobs_end_date, 8, 9, 49
											IF jobs_end_date <> "__ __ __" THEN
												transmit
												EMReadScreen valid_command, 21, 24, 2
											END IF
										LOOP UNTIL valid_command = "ENTER A VALID COMMAND" OR jobs_end_date = "__ __ __"
									END IF
									IF (jobs_info <> "0" AND jobs_end_date = "__ __ __") THEN panel_status = "JOBS"
								END IF
							END IF
						END IF
					END IF
				END IF
			END IF
	NEXT
END FUNCTION

FUNCTION navigate_to_MMIS
	attn

	Do
		EMReadScreen MAI_check, 3, 1, 33
		If MAI_check <> "MAI" then EMWaitReady 1, 1
	Loop until MAI_check = "MAI"

	EMReadScreen mmis_check, 7, 15, 15
	IF mmis_check = "RUNNING" THEN
		EMWriteScreen "10", 2, 15
		transmit
	ELSE
		EMConnect"A"
		attn
		EMReadScreen mmis_check, 7, 15, 15
		IF mmis_check = "RUNNING" THEN
			EMWriteScreen "10", 2, 15
			transmit
		ELSE
			EMConnect"B"
			attn
			EMReadScreen mmis_b_check, 7, 15, 15
			IF mmis_b_check <> "RUNNING" THEN
				script_end_procedure("You do not appear to have MMIS running. This script will now stop. Please make sure you have an active version of MMIS and re-run the script.")
			ELSE
				EMWriteScreen "10", 2, 15
				transmit
			END IF
		END IF
	END IF

	DO
		PF6
		EMReadScreen password_prompt, 38, 2, 23
		IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript
		EMReadScreen session_start, 18, 1, 7
	LOOP UNTIL session_start = "SESSION TERMINATED"

	'Getting back in to MMIS and trasmitting past the warning screen (workers should already have accepted the warning when they logged themselves into MMIS the first time, yo.
	EMWriteScreen "MW00", 1, 2
	transmit
	transmit

	'The following will select the correct version of MMIS. First it looks for C302, then EK01, then C402.
	row = 1
	col = 1
	EMSearch "C302", row, col
	If row <> 0 then 
		If row <> 1 then 'It has to do this in case the worker only has one option (as many LTC and OSA workers don't have the option to decide between MAXIS and MCRE case access). The MMIS screen will show the text, but it's in the first row in these instances.
			EMWriteScreen "x", row, 4
			transmit
		End if
	Else 'Some staff may only have EK01 (MMIS MCRE). The script will allow workers to use that if applicable.
		row = 1
		col = 1
		EMSearch "EK01", row, col
		If row <> 0 then 
			If row <> 1 then
				EMWriteScreen "x", row, 4
				transmit
			End if
		Else 'Some OSAs have C402 (limited access). This will search for that.
			row = 1
			col = 1
			EMSearch "C402", row, col
			If row <> 0 then 
				If row <> 1 then
					EMWriteScreen "x", row, 4
					transmit
				End if
			Else 'Some OSAs have EKIQ (limited MCRE access). This will search for that.
				row = 1
				col = 1
				EMSearch "EKIQ", row, col
				If row <> 0 then 
					If row <> 1 then
						EMWriteScreen "x", row, 4
						transmit
					End if
				Else
					script_end_procedure("C402, C302, EKIQ, or EK01 not found. Your access to MMIS may be limited. Contact your script Alpha user if you have questions about using this script.")
				End if
			End if
		End if
	END IF

	'Now it finds the recipient file application feature and selects it.
	row = 1
	col = 1
	EMSearch "RECIPIENT FILE APPLICATION", row, col
	EMWriteScreen "x", row, col - 3
	transmit
END FUNCTION

FUNCTION navigate_to_MAXIS(maxis_mode)
	attn
	EMConnect "A"
	IF maxis_mode = "PRODUCTION" THEN
		EMReadScreen prod_running, 7, 6, 15
		IF prod_running = "RUNNING" THEN
			x = "A"
		ELSE
			EMConnect"B"
			attn
			EMReadScreen prod_running, 7, 6, 15
			IF prod_running = "RUNNING" THEN
				x = "B"
			ELSE
				script_end_procedure("Please do not run this script in a session larger than 2.")
			END IF
		END IF
	ELSEIF maxis_mode = "INQUIRY DB" THEN
		EMReadScreen inq_running, 7, 7, 15
		IF inq_running = "RUNNING" THEN
			x = "A"
		ELSE
			EMConnect "B"
			attn
			EMReadScreen inq_running, 7, 7, 15
			IF inq_running = "RUNNING" THEN
				x = "B"
			ELSE
				script_end_procedure("Please do not run this script in a session larger than 2.")
			END IF
		END IF
	END IF

	EMConnect (x)
	IF maxis_mode = "PRODUCTION" THEN
		EMWriteScreen "1", 2, 15
		transmit
	ELSEIF maxis_mode = "INQUIRY DB" THEN
		EMWriteScreen "2", 2, 15
		transmit
	END IF		

END FUNCTION

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog pull_cases_into_excel_dialog, 0, 0, 416, 180, "Pull cases into Excel dialog"
  CheckBox 10, 20, 55, 10, "PREG exists?", preg_check
  CheckBox 10, 35, 90, 10, "All HH membs 19+?", all_HH_membs_19_plus_check
  CheckBox 10, 50, 90, 10, "Number of HH membs?", number_of_HH_membs_check
  CheckBox 10, 65, 90, 10, "ABAWD code", ABAWD_code_check
  CheckBox 10, 80, 80, 10, "PDED/Rep-Payee", pded_check
  CheckBox 10, 95, 95, 10, "MAGI%", magi_pct_check
  CheckBox 10, 110, 85, 10, "FS and MFIP Review", FS_MF_review_check
  CheckBox 10, 125, 95, 10, "Homeless Clients", homeless_check
  CheckBox 10, 140, 105, 10, "MAEPD/Part B Reimbursable", maepd_check
  CheckBox 10, 155, 70, 10, "All cases", all_cases_check
  DropListBox 180, 15, 95, 10, "REPT/PND2"+chr(9)+"REPT/ACTV", screen_to_use
  EditBox 190, 30, 90, 15, x_number
  CheckBox 125, 50, 295, 15, "Check here if you're running this for all staff (WARNING: this could take several hours)", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 365, 10, 50, 15
    CancelButton 365, 30, 50, 15
  GroupBox 5, 5, 115, 165, "Additional items to log"
  Text 125, 15, 50, 10, "Screen to use:"
  Text 125, 35, 60, 10, "Worker to check:"
EndDialog

BeginDialog gen_worker_dialog, 0, 0, 291, 110, "Pull cases into Excel dialog"
  CheckBox 10, 20, 55, 10, "PREG exists?", preg_check
  CheckBox 10, 35, 90, 10, "ABAWD code", ABAWD_code_check
  CheckBox 10, 50, 80, 10, "PDED/Rep-Payee", pded_check
  CheckBox 10, 65, 85, 10, "Homeless Clients", homeless_check
  CheckBox 10, 80, 105, 10, "MA-EPD/Part B Reimbursable", maepd_check
  DropListBox 185, 15, 95, 10, "REPT/PND2"+chr(9)+"REPT/ACTV", screen_to_use
  ButtonGroup ButtonPressed
    OkButton 175, 50, 50, 15
    CancelButton 230, 50, 50, 15
  Text 130, 15, 50, 10, "Screen to use:"
  GroupBox 5, 5, 120, 90, "Additional items to log"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""

'Checking for MAXIS
MAXIS_check_function

'Grabbing user ID to validate user of script. Only some users are allowed to use this script.
Set objNet = CreateObject("WScript.NetWork") 
user_ID_for_validation = ucase(objNet.UserName)

'Validating user ID
If user_ID_for_validation = "VKCARY" OR _
	user_ID_for_validation = "RAKALB" OR _
	user_ID_for_validation = "CDPOTTER" OR _ 
	user_ID_for_validation = "MLDIETZ" OR _ 
	user_ID_for_validation = "PHBROCKM" OR _
	user_ID_for_validation = "JGLETH" OR _
	user_ID_for_validation = "TMMIELKE" OR _ 
	user_ID_for_validation = "VLANDERS" OR _ 
	user_ID_for_validation = "SLCARDA" OR _ 
	user_ID_for_validation = "IGFERRIS" OR _ 
	user_ID_for_validation = "CMCOX" THEN 
	Dialog pull_cases_into_excel_dialog
		If buttonpressed = 0 then stopscript
ELSE
	DIALOG gen_worker_dialog
		IF ButtonPressed = 0 THEN stopscript
END IF

IF x_number = "" THEN CALL find_variable("User: ", x_number, 7)

'Adjusting name of script variable for usage stats according to what was done. So, if ACTV was used instead of PND2, it'll indicate that on the script (and thus allow accurate measurement of time savings).
If screen_to_use = "REPT/PND2" then
	name_of_script = "BULK - pull cases into Excel (PND2)"
	If all_workers_check = 1 then name_of_script = "BULK - pull cases into Excel (PND2 all cases)"
ElseIf screen_to_use = "REPT/ACTV" then
	name_of_script = "BULK - pull cases into Excel (ACTV)"
	If all_workers_check = 1 then name_of_script = "BULK - pull cases into Excel (ACTV all cases)"
End if


'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True


'Setting the first 3 col as worker, case number, and name
ObjExcel.Cells(1, 1).Value = "X Number"
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
ObjExcel.Cells(1, 3).Value = "NAME"

'If working off of PND2 it sets the 4th  col as APPL DATE, otherwise it'll be NEXT REVW DATE
If screen_to_use = "REPT/PND2" then
	ObjExcel.Cells(1, 4).Value = "APPL DATE"
ElseIf screen_to_use = "REPT/ACTV" then
	ObjExcel.Cells(1, 4).Value = "NEXT REVW DATE"	
End if

'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
'	Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 5 'Starting with 4 because cols 1-3 are already used
If preg_check = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "PREG EXISTS?"
	preg_col = col_to_use
	col_to_use = col_to_use + 1
End if
If all_HH_membs_19_plus_check = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "ALL MEMBS 19+?"
	all_HH_membs_19_plus_col = col_to_use
	col_to_use = col_to_use + 1
End if
If number_of_HH_membs_check = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "NUMBER OF HH MEMBS?"
	number_of_HH_membs_col = col_to_use
	col_to_use = col_to_use + 1
End if
If ABAWD_code_check = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "ABAWD CODE"
	ABAWD_code_col = col_to_use
	col_to_use = col_to_use + 1
End if
IF pded_check = 1 THEN
	ObjExcel.Cells(1, col_to_use).Value = "PDED/Rep-Payee"
	pded_col = col_to_use
	col_to_use = col_to_use + 1
END IF
IF magi_pct_check = 1 THEN
	ObjExcel.Cells(1, col_to_use).Value = "All MAGI HH"
	magi_col = col_to_use
	col_to_use = col_to_use + 1
END IF
IF FS_MF_review_check = 1 THEN
	ObjExcel.Cells(1, col_to_use).Value = "SNAP Cases to Review"
	SNAP_col = col_to_use
	col_to_use = col_to_use + 1
	ObjExcel.Cells(1, col_to_use).Value = "MFIP Cases to Review"
	MFIP_col = col_to_use
	col_to_use = col_to_use + 1
END IF
IF homeless_check = 1 THEN
	objExcel.Cells(1, col_to_use).Value = "CL reporting homeless?"
	homeless_col = col_to_use
	col_to_use = col_to_use + 1
END IF
IF maepd_check = 1 THEN
	objExcel.Cells(1, col_to_use).Value = "MA-EPD & Part B Reimburseable"
	maepd_col = col_to_use
	col_to_use = col_to_use + 1
END IF
IF all_cases_check = 1 THEN
	screen_to_use = "REPT/ACTV"
	all_workers_check = 1
END IF



'Setting the variable for what's to come
excel_row = 2

'If all workers are selected, the script will open the worker list stored on the shared drive, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
IF magi_pct_check = 1 THEN
	x_array = "293" & " " & "692" & " " & "107" & " " & "B98" & " " & "SAC" & " " & "234" & " " & "30V" & " " & "757" & " " & "628" & " " & "4AF" & " " & "144" & " " & "880" & " " & "B83" & " " & "130" & " " & "769" & " " & "524" & " " & "132" & " " & "598" & " " & "752" & " " & "4SZ" & " " & "4SS" & " " & "B93" & " " & "950" & " " & "742" & " " & "112" & " " & "756" & " " & "TRP" & " " & "111" & " " & "EMP" & " " & "SKM" & " " & "268" & " " & "722" & " " & "B42" & " " & "4BL" & " " & "233" & " " & "122" & " " & "932" & " " & "894" & " " & "770" & " " & "976" & " " & "989" & " " & "949" & " " & "631" & " " & "978" & " " & "5A2" & " " & "104" & " " & "4SY" & " " & "GMZ"
	x_array = split(x_array)
ELSE
	If all_workers_check = 1 then
		CALL create_array_of_all_active_x_numbers_in_county(x_array, "02")
	Else
		IF len(x_number) > 3 THEN 
			x_array = split(x_number, ", ")
		ELSE		
			x_array = split(x_number)
		END IF
	End if
END IF

For each worker in x_array
'Getting to PND2, if PND2 is the selected option
If screen_to_use = "REPT/PND2" then
	Call navigate_to_screen("rept", "pnd2")
	EMWriteScreen worker, 21, 13
	transmit

	'Grabbing each case number on screen
	Do
		MAXIS_row = 7
		Do
			EMReadScreen case_number, 8, MAXIS_row, 5
			If case_number = "        " then 
				EMReadScreen additional_app, 14, maxis_row, 17
				IF additional_app = "              " THEN
					EXIT DO
				ELSE
					MAXIS_row = MAXIS_row + 1
				END IF
			ELSE
				EMReadScreen client_name, 22, MAXIS_row, 16
				EMReadScreen APPL_date, 8, MAXIS_row, 38
				ObjExcel.Cells(excel_row, 1).Value = worker
				ObjExcel.Cells(excel_row, 2).Value = case_number
				ObjExcel.Cells(excel_row, 3).Value = client_name
				ObjExcel.Cells(excel_row, 4).Value = replace(APPL_date, " ", "/")
				MAXIS_row = MAXIS_row + 1
			END IF
			excel_row = excel_row + 1
		Loop until MAXIS_row = 19
		PF8
		EMReadScreen last_page_check, 21, 24, 2
	Loop until last_page_check = "THIS IS THE LAST PAGE"
End if

'Getting to ACTV, if ACTV is the selected option
If screen_to_use = "REPT/ACTV" then
	Call navigate_to_screen("rept", "actv")
	IF worker <> "" THEN
		EMWriteScreen worker, 21, 13
		transmit
	END IF
	EMReadScreen user_id, 7, 21, 71
	EMReadScreen check_worker, 7, 21, 13
	IF user_id = check_worker THEN PF7

	'Grabbing each case number on screen
	Do
		MAXIS_row = 7
		EMReadScreen last_page_check, 21, 24, 2
		Do
			EMReadScreen case_number, 8, MAXIS_row, 12
			If case_number = "        " then exit do
			EMReadScreen client_name, 21, MAXIS_row, 21
			EMReadScreen next_REVW_date, 8, MAXIS_row, 42
			ObjExcel.Cells(excel_row, 1).Value = worker
			ObjExcel.Cells(excel_row, 2).Value = case_number
			ObjExcel.Cells(excel_row, 3).Value = client_name
			ObjExcel.Cells(excel_row, 4).Value = replace(next_REVW_date, " ", "/")
			MAXIS_row = MAXIS_row + 1
			excel_row = excel_row + 1
		Loop until MAXIS_row = 19
		PF8
	Loop until last_page_check = "THIS IS THE LAST PAGE"
End if

next

'Resetting excel_row variable, now we need to start looking people up
excel_row = 2 

Do 
	case_number = ObjExcel.Cells(excel_row, 2).Value
	If case_number = "" then exit do

	'Now pulling PREG info
	If preg_check = 1 then
		call navigate_to_screen("STAT", "PREG")
		EMReadScreen PREG_panel_check, 1, 2, 78
		If PREG_panel_check <> "0" then 
			ObjExcel.Cells(excel_row, preg_col).Value = "Y"
		Else
			ObjExcel.Cells(excel_row, preg_col).Value = "N"
		End if
	End if

	'Now pulling age info
	If all_HH_membs_19_plus_check = 1 then
		call navigate_to_screen("STAT", "MEMB")
		Do
			EMReadScreen MEMB_panel_current, 1, 2, 73
			EMReadScreen MEMB_panel_total, 1, 2, 78
			EMReadScreen MEMB_age, 3, 8, 76
			If MEMB_age = "   " then MEMB_age = "0"
			If cint(MEMB_age) < 19 then has_minor_in_case = True
			transmit
		Loop until MEMB_panel_current = MEMB_panel_total
		If has_minor_in_case <> True then 
			ObjExcel.Cells(excel_row, all_HH_membs_19_plus_col).Value = "Y"
		Else
			ObjExcel.Cells(excel_row, all_HH_membs_19_plus_col).Value = "N"
		End if
		has_minor_in_case = "" 'clearing variable
	End if

	'Now pulling number of membs info
	If number_of_HH_membs_check = 1 then
		call navigate_to_screen("STAT", "MEMB")
		EMReadScreen MEMB_panel_total, 1, 2, 78
		ObjExcel.Cells(excel_row, number_of_HH_membs_col).Value = cint(MEMB_panel_total)
	End if

	'Now pulling ABAWD info
	If ABAWD_code_check = 1 then
		ABAWD_status = "" 		'clearing variable
		eats_group_members = ""		'clearing

		call navigate_to_screen("STAT", "PROG")
		ERRR_screen_check
		
		EMReadScreen snap_status, 4, 10, 74
		IF snap_status = "ACTV" OR snap_status = "PEND" THEN
			call navigate_to_screen("STAT", "EATS")
			ERRR_screen_check
			EMReadScreen all_eat_together, 1, 4, 72
			IF all_eat_together = "_" THEN
				eats_group_members = "01" & " "
			ELSEIF all_eat_together = "Y" THEN 
				eats_row = 5
				DO
					EMReadScreen eats_person, 2, eats_row, 3
					eats_person = replace(eats_person, " ", "")
					IF eats_person <> "" THEN 
						eats_group_members = eats_group_members & eats_person & " "
						eats_row = eats_row + 1
					END IF
				LOOP UNTIL eats_person = ""
			ELSEIF all_eat_together = "N" THEN
				eats_row = 13
				DO
					EMReadScreen eats_group, 38, eats_row, 39
					find_memb01 = InStr(eats_group, "01")
					IF find_memb01 = 0 THEN eats_row = eats_row + 1
				LOOP UNTIL find_memb01 <> 0
				eats_col = 39
				DO
					EMReadScreen eats_group, 2, eats_row, eats_col
					IF eats_group <> "__" THEN 
						eats_group_members = eats_group_members & eats_group & " "
						eats_col = eats_col + 4
					END IF
				LOOP UNTIL eats_group = "__"
			END IF

			eats_group_members = trim(eats_group_members)
			eats_group_members = split(eats_group_members)

			call navigate_to_screen("STAT", "WREG")
			ERRR_screen_check
	
			FOR EACH person IN eats_group_members
				EMWriteScreen person, 20, 76
				transmit
				
				EMReadScreen ABAWD_status_code, 2, 13, 50
				ABAWD_status = ABAWD_status & person & ": " & ABAWD_status_code & ","
			NEXT
	
			ObjExcel.Cells(excel_row, ABAWD_code_col).Value = ABAWD_status

		End if

		IF objExcel.Cells(excel_row, ABAWD_code_col).Value = "" THEN 
			SET objRange = objExcel.Cells(excel_row, 1).EntireRow
			objRange.Delete
			excel_row = excel_row - 1
		End IF
	End if

	IF pded_check = 1 THEN
		total_pded = ""
		pded_hh_array = ""
		call navigate_to_screen("STAT", "PDED")
			errr_screen_check
			pded_row = 5
			DO
				EMReadScreen pded_hh_memb, 2, pded_row, 3
				IF pded_hh_memb = "  " THEN
					EXIT DO
				ELSE
					pded_hh_array = pded_hh_array & pded_hh_memb & " "
					pded_row = pded_row + 1
				END IF
			LOOP UNTIL pded_hh_memb = "  "

			pded_hh_array = trim(pded_hh_array)
			pded_hh_array = split(pded_hh_array)

			FOR EACH hh_memb IN pded_hh_array
				pded_info = ""
				rep_payee_amt = ""
				EMWriteScreen hh_memb, 20, 76
				transmit
					EMReadScreen rep_payee_amt, 8, 15, 70
				rep_payee_amt = replace(rep_payee_amt, "_", "")
				rep_payee_amt = replace(rep_payee_amt, " ", "")
				IF rep_payee_amt <> "" THEN
					pded_info = hh_memb & ": " & rep_payee_amt & "; "
					total_pded = total_pded & pded_info
				END IF
			NEXT

			ObjExcel.Cells(excel_row, PDED_col).Value = total_pded
			'THE FOLLOWING 5 LINES AUTOMATICALLY DELETE ANY BLANK RESULTS
			IF total_pded = "" THEN
				Set objRange = objExcel.Cells(excel_row, 1).EntireRow
				objRange.Delete
				excel_row = excel_row - 1
			END IF

	END IF

	IF magi_pct_check = 1 THEN
		'	'Finds HC Budget Method
		MAGI_status = ""
		MAGI_result = ""
		call navigate_to_screen("ELIG", "HC")
		hhmm_row = 8
		DO
			EMReadScreen hc_ref_num, 2, hhmm_row, 3
			IF hc_ref_num <> "  " THEN
				EMReadScreen hc_requested, 1, hhmm_row, 28
				IF hc_requested = "S" OR hc_requested = "Q" OR hc_requested = "I" THEN 			'IF the HH MEMB is MSP ONLY then they are automatically Budg Mthd B
					MAGI_status = MAGI_status & "B"
					hhmm_row = hhmm_row + 1
				ELSEIF hc_requested = "M" THEN
					EMWriteScreen "X", hhmm_row, 26
					transmit
					EMReadScreen budg_mthd, 1, 13, 76
					MAGI_status = MAGI_status & budg_mthd
					PF3
					hhmm_row = hhmm_row + 1
				ELSEIF hc_requested = "N" OR hc_requested = "E" THEN
					hhmm_row = hhmm_row + 1
				END IF
			ELSEIF hc_ref_num = "  " THEN
				EXIT DO			
			END IF
		LOOP UNTIL hhmm_row = 20 OR hc_ref_num = "  "
		IF MAGI_status <> "" AND ((InStr(MAGI_status, "B") <> 0) OR (InStr(MAGI_status, "X") <> 0)) THEN
			MAGI_result = "Non-MAGI"
		ELSEIF MAGI_status <> "" AND ((InStr(MAGI_status, "B") = 0) AND (InStr(MAGI_status, "X") = 0)) THEN
			MAGI_result = "MAGI"
		ELSEIF MAGI_status = "" THEN
			MAGI_result = ""
		END IF
		ObjExcel.Cells(excel_row, magi_col).Value = MAGI_result
		IF MAGI_result = "" THEN 
			SET objRange = objExcel.Cells(excel_row, 1).EntireRow
			objRange.Delete
			excel_row = excel_row - 1
		END IF
	END IF 

	IF FS_MF_review_check = 1 THEN
		panel_status = ""
		snap_panel_status = ""
		cash_panel_status = ""
		eats_group_members = ""
		mfip_group = ""
		CALL navigate_to_screen("STAT", "PROG")
		ERRR_screen_check

		EMReadScreen snap_status, 4, 10, 74
		EMReadScreen MFIP_prog_1, 2, 6, 67
		EMReadScreen MFIP_prog_2, 2, 7, 67
		EMReadScreen MFIP_status_1, 4, 6, 74
		EMReadScreen MFIP_status_2, 4, 7, 74

		IF snap_status = "ACTV" OR (MFIP_prog_1 = "MF" AND MFIP_status_1 = "ACTV") OR (MFIP_prog_2 = "MF" AND MFIP_status_2 = "ACTV") THEN

			CALL navigate_to_screen("STAT", "REVW")
			ERRR_screen_check			
			EmReadScreen cash_review_date, 8, 09, 37   'reads cash renewal date
			EMwritescreen "X", 5, 58
			Transmit
			EmReadScreen snap_review_date, 8, 09, 64     'reads snap ER date
			cash_review_date = replace(cash_review_date, " ", "/")
			snap_review_date = replace(snap_review_date, " ", "/")

			comparison_date = datepart("M", date) & "/01/" & datepart("yyyy", date)
			past_month = dateadd("M", -1, comparison_date)				'establishes minimum range
			future_month = dateadd("M", 1, comparison_date)				'establishes maximum range

			'calcuate the past current and future renewal months and years
			review_date_1_year_past = dateadd("YYYY", 1, past_month)
			review_date_1_year_current = dateadd("YYYY", 1, comparison_date)
			review_date_1_year_future = dateadd("YYYY", 1, future_month)
			review_date_2_year_past = dateadd("YYYY", 2, past_month)
			review_date_2_year_current = dateadd("YYYY", 2, comparison_date)	
			review_date_2_year_future = dateadd("YYYY", 2, future_month)

			IF snap_status = "ACTV" THEN
				IF (cdate(snap_review_date) = cdate(review_date_1_year_past) OR _
					cdate(snap_review_date) = cdate(review_date_1_year_current) OR _
					cdate(snap_review_date) = cdate(review_date_1_year_future) OR _
					cdate(snap_review_date) = cdate(review_date_2_year_past) OR _
					cdate(snap_review_date) = cdate(review_date_2_year_current) OR _
					cdate(snap_review_date) = cdate(review_date_2_year_future)) THEN

					call navigate_to_screen("STAT", "EATS")
					EMReadScreen all_eat_together, 1, 4, 72
					IF all_eat_together = "_" THEN
						eats_group_members = "01" & " "
					ELSEIF all_eat_together = "Y" THEN 
						eats_row = 5
						DO
							EMReadScreen eats_person, 2, eats_row, 3
							eats_person = replace(eats_person, " ", "")
							IF eats_person <> "" THEN 
								eats_group_members = eats_group_members & eats_person & " "
								eats_row = eats_row + 1
							END IF
						LOOP UNTIL eats_person = ""
					ELSEIF all_eat_together = "N" THEN
						eats_row = 13
						DO
							EMReadScreen eats_group, 38, eats_row, 39
							find_memb01 = InStr(eats_group, "01")
							IF find_memb01 = 0 THEN eats_row = eats_row + 1
						LOOP UNTIL find_memb01 <> 0
						eats_col = 39
						DO
							EMReadScreen eats_group, 2, eats_row, eats_col
							IF eats_group <> "__" THEN 
								eats_group_members = eats_group_members & eats_group & " "
								eats_col = eats_col + 4
							END IF
						LOOP UNTIL eats_group = "__"
					END IF

					eats_group_members = trim(eats_group_members)
					eats_group_members = split(eats_group_members)
			
					CALL check_panels_function(eats_group_members, panel_status)
					snap_panel_status = panel_status
					IF snap_panel_status <> "" THEN ObjExcel.Cells(excel_row, snap_col).Value = "Review SNAP"
				END IF
			END IF	
			IF ((MFIP_prog_1 = "MF" AND MFIP_status_1 = "ACTV") OR (MFIP_prog_2 = "MF" AND MFIP_status_2 = "ACTV")) THEN 
				IF (cdate(cash_review_date) = cdate(review_date_1_year_past) OR _
					cdate(cash_review_date) = cdate(review_date_1_year_current) OR _
					cdate(cash_review_date) = cdate(review_date_1_year_future) OR _
					cdate(cash_review_date) = cdate(review_date_2_year_past) OR _
					cdate(cash_review_date) = cdate(review_date_2_year_current) OR _
					cdate(cash_review_date) = cdate(review_date_2_year_future)) THEN
					
					panel_status = ""
				
					CALL navigate_to_screen("ELIG", "MFIP")
					mfpr_row = 7
					DO
						IF mfpr_row = 18 THEN 
							PF8
							EMReadScreen no_more_members, 15, 24, 5
							mfpr_row = 7
						END IF
						EMReadScreen is_counted, 7, mfpr_row, 41
						is_counted = replace(is_counted, " ", "")
						IF is_counted = "COUNTED" THEN 
							EMReadScreen ref_num, 2, mfpr_row, 6
							mfip_group = mfip_group & ref_num & " "
						END IF
						mfpr_row = mfpr_row + 1
					LOOP UNTIL is_counted = "" OR no_more_members = "NO MORE MEMBERS"
					mfip_group = trim(mfip_group)
					mfip_group = split(mfip_group)
	
					CALL check_panels_function(mfip_group, panel_status)
					cash_panel_status = panel_status
					IF cash_panel_status <> "" THEN ObjExcel.Cells(excel_row, mfip_col).Value = "Review MFIP"
				END IF
			END IF
		END IF

		IF snap_panel_status = "" AND cash_panel_status = "" THEN 
			SET objRange = objExcel.Cells(excel_row, 1).EntireRow
			objRange.Delete
			excel_row = excel_row - 1
		End IF
	END IF

	IF homeless_check = 1 THEN
		CALL navigate_to_screen("STAT", "ADDR")
		ERRR_screen_check
		EMReadScreen addr_line, 16, 6, 43
		EMReadScreen homeless_yn, 1, 10, 43
		IF homeless_yn = "Y" OR addr_line = "GENERAL DELIVERY" THEN 
			objExcel.Cells(excel_row, homeless_col).Value = "HOMELESS"
		ELSEIF homeless_yn <> "Y" AND addr_line <> "GENERAL DELIVERY" THEN
			SET objRange = objExcel.Cells(excel_row, 1).EntireRow
			objRange.Delete
			excel_row = excel_row - 1
		END IF			
	END IF

	IF MAEPD_check = 1 THEN
		back_to_SELF
		CALL find_variable("Environment: ", production_or_inquiry, 10)
		CALL navigate_to_screen("ELIG", "HC")
		hhmm_row = 8
		DO
			EMReadScreen hc_type, 2, hhmm_row, 28
			IF hc_type = "MA" THEN
				EMWriteScreen "X", hhmm_row, 26
				transmit
				EMReadScreen elig_type, 2, 12, 72
				IF elig_type = "DP" THEN
					EMWriteScreen "X", 9, 76
					transmit
					EMReadScreen pct_fpg, 4, 18, 38
					pct_fpg = trim(pct_fpg)
					pct_fpg = pct_fpg * 1
					IF pct_fpg < 201 THEN
						PF3
						PF3
						EMReadScreen hh_memb_num, 2, hhmm_row, 3
						CALL navigate_to_screen("STAT", "MEMB")
						ERRR_screen_check
						EMWriteScreen hh_memb_num, 20, 76
						transmit
						EMReadScreen cl_pmi, 8, 4, 46
						cl_pmi = replace(cl_pmi, " ", "")
						DO
							IF len(cl_pmi) <> 8 THEN cl_pmi = "0" & cl_pmi
						LOOP UNTIL len(cl_pmi) = 8
						navigate_to_MMIS
						DO
							EMReadScreen RKEY, 4, 1, 52
							IF RKEY <> "RKEY" THEN EMWaitReady 0, 0
						LOOP UNTIL RKEY = "RKEY"
						EMWriteScreen "I", 2, 19
						EMWriteScreen cl_pmi, 4, 19
						transmit
						EMWriteScreen "RELG", 1, 8
						transmit
				
						'Reading RELG to determine if the CL is active on MA-EPD		
						EMReadScreen prog01_type, 8, 6, 13
							EMReadScreen elig01_type, 2, 6, 33
							EMReadScreen elig01_end, 8, 7, 36
						EMReadScreen prog02_type, 8, 10, 13
							EMReadScreen elig02_type, 2, 10, 33
							EMReadScreen elig02_end, 8, 11, 36
						EMReadScreen prog03_type, 8, 14, 13
							EMReadScreen elig03_type, 2, 14, 33
							EMReadScreen elig03_end, 8, 15, 36
						EMReadScreen prog04_type, 8, 18, 13
							EMReadScreen elig04_type, 2, 18, 33
							EMReadScreen elig04_end, 8, 19, 36

						IF ((prog01_type = "MEDICAID" AND elig01_type = "DP" AND elig01_end = "99/99/99") OR _
							(prog02_type = "MEDICAID" AND elig02_type = "DP" AND elig02_end = "99/99/99") OR _
							(prog03_type = "MEDICAID" AND elig03_type = "DP" AND elig03_end = "99/99/99") OR _
							(prog04_type = "MEDICAID" AND elig04_type = "DP" AND elig04_end = "99/99/99")) THEN
				
							EMWriteScreen "RMCR", 1, 8
							transmit

							'-----CHECKING FOR ON-GOING MEDICARE PART B-----
							EMReadScreen part_b_begin01, 8, 13, 4
								part_b_begin01 = trim(part_b_begin01)
							EMReadScreen part_b_end01, 8, 13, 15
							EMReadScreen part_b_begin02, 8, 14, 4
								part_b_begin02 = trim(part_b_begin02)
							EMReadScreen part_b_end02, 8, 14, 15
							
							IF (part_b_begin01 <> "" AND part_b_end01 = "99/99/99") THEN		
								EMWriteScreen "RBYB", 1, 8
								transmit
								
								EMReadScreen accrete_date, 8, 5, 66
								EMReadScreen delete_date, 8, 6, 65
								accrete_date = replace(accrete_date, " ", "")

								IF ((accrete_date = "") OR (accrete_date <> "" AND delete_date <> "99/99/99")) THEN
									objExcel.Cells(excel_row, maepd_col).Value = objExcel.Cells(excel_row, maepd_col).Value & ("MEMB " & hh_memb_num & " ELIG FOR REIMBURSEMENT, ")
								END IF
								PF3
							END IF
						ELSE
							PF3
						END IF
						CALL navigate_to_MAXIS(production_or_inquiry)
						hhmm_row = hhmm_row + 1
						CALL navigate_to_screen("ELIG", "HC")
					ELSE
						DO
							EMReadScreen at_hhmm, 4, 3, 51
							IF at_hhmm <> "HHMM" THEN PF3
						LOOP UNTIL at_hhmm = "HHMM"
						hhmm_row = hhmm_row + 1
					END IF
				ELSE
					PF3
					hhmm_row = hhmm_row + 1
				END IF
			ELSE
				hhmm_row = hhmm_row + 1
			END IF
			IF hhmm_row = 20 THEN
				PF8
				EMReadScreen this_is_the_last_page, 21, 24, 2
			END IF
		LOOP UNTIL hc_type = "  " OR this_is_the_last_page = "THIS IS THE LAST PAGE"
		'Deleting the blank results to clean up the spreadsheet
		IF objExcel.Cells(excel_row, maepd_col).Value = "" THEN
			SET objRange = objExcel.Cells(excel_row, 1).EntireRow
			objRange.Delete
			excel_row = excel_row - 1
		END IF				
	END IF

	excel_row = excel_row + 1
Loop until case_number = ""

IF magi_pct_check = 1 THEN
	objExcel.Cells(excel_row + 2, magi_col - 1).Value = "Pct MAGI"
	objExcel.Cells(excel_row + 2, magi_col).Value = "=countif(E2:E" & (excel_row - 1) & ", " & Chr(34) & "MAGI" & Chr(34) & ")/" & (excel_row - 2)
END IF

'Logging usage stats
script_end_procedure("DONE!!")
