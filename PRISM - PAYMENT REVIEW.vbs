DIM ENFL_case_based
DIM ENFL_person_based
DIM CAAD_nav_button
DIM CAHL_nav_button
DIM MAXIS_nav_button
DIM MMIS_nav_button
DIM review_result
DIM current_dlg
DIM page_count
DIM ButtonPressed
DIM CAFS_button
DIM ENFL_button
DIM INWD_button
DIM LETL_button
DIM NCDD_button
DIM NCID_button
DIM PALC_button
DIM PAPD_button
DIM SUDL_button
DIM CAFS_nav_button
DIM ENFL_nav_button
DIM INWD_nav_button
DIM LETL_nav_button
DIM NCDD_nav_button
DIM NCID_nav_button
DIM PALC_nav_button
DIM PAPD_nav_button
DIM SUDL_nav_button
DIM PRISM_case_number
DIM total_arrears
DIM npa_arrears
DIM pa_arrears
DIM mo_non_acc
DIM monthly_accrual
DIM addr_date
DIM addr_known
DIM SUDL
DIM SUDL_display
DIM LETL
DIM LETL_display
DIM ENFL

Function attn
  EMSendKey "<attn>"
  EMWaitReady -1, 0
End function

Function find_variable(x, y, z) 'x is string, y is variable, z is length of new variable
  row = 1
  col = 1
  EMSearch x, row, col
  If row <> 0 then EMReadScreen y, z, row, col + len(x)
End function

Function navigate_to_PRISM_screen(x) 'x is the name of the screen
  EMWriteScreen x, 21, 18
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End function

Function PF1
  EMSendKey "<PF1>"
  EMWaitReady 0, 0
End function

Function PF2
  EMSendKey "<PF2>"
  EMWaitReady 0, 0
End function

function PF3
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
end function

Function PF4
  EMSendKey "<PF4>"
  EMWaitReady 0, 0
End function

Function PF5
  EMSendKey "<PF5>"
  EMWaitReady 0, 0
End function

Function PF6
  EMSendKey "<PF6>"
  EMWaitReady 0, 0
End function

Function PF7
  EMSendKey "<PF7>"
  EMWaitReady 0, 0
End function

function PF8
  EMSendKey "<PF8>"
  EMWaitReady 0, 0
end function

function PF9
  EMSendKey "<PF9>"
  EMWaitReady 0, 0
end function

function PF10
  EMSendKey "<PF10>"
  EMWaitReady 0, 0
end function

Function PF11
  EMSendKey "<PF11>"
  EMWaitReady 0, 0
End function

Function PF12
  EMSendKey "<PF12>"
  EMWaitReady 0, 0
End function

function PF20
  EMSendKey "<PF20>"
  EMWaitReady 0, 0
end function

function transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
end function

'-----NEW FUNCTIONS-----
'-----Navigation-----
FUNCTION go_to(input_value, row, col)
	EMWriteScreen input_value, row, col
	transmit
END FUNCTION

'----CAFS----
FUNCTION read_monthly_non_accrual(monthly_non_accrual_variable)
	CALL navigate_to_prism_screen("CAFS")
	CALL find_variable("Monthly Nonaccrual    :", monthly_non_accrual_variable, 14)
	monthly_non_accrual_variable = trim(monthly_non_accrual_variable)
END FUNCTION

FUNCTION read_monthly_accrual(monthly_accrual_variable)
	CALL navigate_to_prism_screen("CAFS")
	CALL find_variable("Monthly Accrual       :", monthly_accrual_variable, 15)		'Commented out because we are not sure why it started reading "PR"
	monthly_accrual_variable = trim(monthly_accrual_variable)
END FUNCTION

FUNCTION read_unpaid_monthly_accrual(unpaid_monthly_accrual_variable)
	CALL navigate_to_prism_screen("CAFS")
	CALL find_variable("Unpaid Monthly Accrual:", unpaid_monthly_accrual_variable, 14)
	unpaid_monthly_accrual_variable = trim(unpaid_monthly_accrual_variable)
END FUNCTION

FUNCTION read_unpaid_monthly_non_accrual(unpaid_monthly_non_accrual_variable)
	CALL navigate_to_prism_screen("CAFS")
	CALL find_variable("Unpaid Mo Non-Accrual :", unpaid_monthly_non_accrual_variable, 14)
	unpaid_monthly_non_accrual_variable = trim(unpaid_monthly_non_accrual_variable)
END FUNCTION

FUNCTION read_cafs_past_due(past_due_variable)
	CALL navigate_to_prism_screen("CAFS")
	CALL find_variable("Past Due              :", past_due_variable, 14)
	past_due_variable = trim(past_due_variable)
END FUNCTION

FUNCTION read_cafs_total_due(total_due_variable)
	CALL navigate_to_prism_screen("CAFS")
	CALL find_variable("Total Due             :", total_due_variable, 14)
	total_due_variable = trim(total_due_variable)
END FUNCTION

FUNCTION read_cafs_suspense(suspense_variable)
	CALL navigate_to_prism_screen("CAFS")
	CALL find_variable("Total Due             :", suspense_variable, 14)
	suspense_variable = trim(suspense_variable)
END FUNCTION

FUNCTION read_npa_arrears(npa_arrears_variable)
	CALL navigate_to_prism_screen("CAFS")
	CALL find_variable("NPA Arrears    :", npa_arrears_variable, 14)
	npa_arrears_variable = trim(npa_arrears_variable)
END FUNCTION

FUNCTION read_pa_arrears(pa_arrears_variable)
	CALL navigate_to_prism_screen("CAFS")
	CALL find_variable("PA Arrears     :", pa_arrears_variable, 14)
	pa_arrears_variable = trim(pa_arrears_variable)
END FUNCTION

FUNCTION read_total_arrears(total_arrears_variable)
	CALL navigate_to_prism_screen("CAFS")
	CALL find_variable("PA Arrears     :", total_arrears_variable, 14)
	total_arrears_variable = trim(total_arrears_variable)
END FUNCTION	

'-----PALC-----
FUNCTION read_PALC_payment_type(payment_type)
	CALL navigate_to_PRISM_screen("PALC")
	EMReadScreen payment_type, 3, 9, 25
END FUNCTION

FUNCTION read_PALC_last_payment_date(last_payment_date)
	CALL navigate_to_PRISM_screen("PALC")
	EMWriteScreen "D", 9, 5
	transmit
	EMReadScreen last_payment_date, 8, 13, 37
	PF3
END FUNCTION

FUNCTION read_PALC_payment_amount(payment_amount)
	CALL navigate_to_PRISM_screen("PALC")
	EMReadScreen payment_amount, 13, 9, 29
	payment_amount = trim(payment_amount)
END FUNCTION

FUNCTION read_PALC_alloc_amount(alloc_amount)
	CALL navigate_to_PRISM_screen("PALC")
	EMReadScreen alloc_amount, 12, 9, 68
	alloc_amount = trim(alloc_amount)
END FUNCTION
	

'-----NCDD----
FUNCTION read_NCDD_address_known(addr_known)
	CALL navigate_to_PRISM_screen("NCDD")
	CALL find_variable("Address Known: ", addr_known, 1)
END FUNCTION

FUNCTION read_NCDD_address_effective_date(addr_date)
	CALL navigate_to_PRISM_screen("NCDD")
	CALL find_variable("Ver: ", addr_date, 10)
END FUNCTION

'-----SUDL-----
FUNCTION create_SUDL_variable(SUDL, sudl_row)
	CALL navigate_to_PRISM_screen("SUDL")
	CALL go_to ("Y", 20, 78)
	
	DO
		EMReadScreen remedy, 3, sudl_row, 9
		remedy = trim(remedy)
		IF remedy <> "" THEN 
			EMReadScreen supp_code, 1, sudl_row, 16
			SUDL = SUDL & "Remedy: " & remedy & ", " & " Suppress Code: " & supp_code	& "; "
		END IF
		sudl_row = sudl_row + 1
	LOOP UNTIL remedy = "" 
END FUNCTION

'-----LETL-----
FUNCTION create_LETL_variable(LETL, letl_row)
	CALL navigate_to_PRISM_screen("LETL")
	
	DO
		EMReadScreen begin_date, 8, letl_row, 10
		begin_date = trim(begin_date)
		IF begin_date <> "" THEN 
			EMReadScreen legal_process, 3, letl_row, 21
			LETL = LETL & "Begin Date: " & begin_date & ", " & "Legal Process: " & legal_process & "; "
		END IF
		letl_row = letl_row + 1
	LOOP UNTIL begin_date = ""
END FUNCTION

'-----INWD-----
FUNCTION create_INWD_array
		'Sets current position of first employer
	employer_array_position = 0
		'Navigate to screen (I assume... Robert you will have to document here)
	CALL navigate_to_PRISM_screen("INWD")
	CALL go_to("B", 3, 29)
	row = 7
	DO
		DO
			'Checks to see if you are at the end of the data
			'EMReadScreen end_of_data, 11, 24, 2
			EMReadScreen end_of_data, 19, row, 28
			'Checks to see if your screen shows an active Job
			EMReadScreen job_status, 3, row, 2
			IF job_status = "ACT" THEN
					'This now only advances the array when it actually gets valid data
				ReDim inwd_array(employer_array_position, 12)
					
				EMSetCursor row, 2
				transmit
					'Filling the variables on INWD array for current job
				'(Ubound(inwd_array,1)) sets the position equal to the current row in the array
				' Array columns start at 0. So 13 values is 0 through 12
				inwd_array((Ubound(inwd_array,1)),0)  = read_INWD_employer_name
				inwd_array((Ubound(inwd_array,1)),1)  = read_INWD_mo_acc_basic_support
				inwd_array((Ubound(inwd_array,1)),2)  = read_INWD_mo_acc_spousal_maint
				inwd_array((Ubound(inwd_array,1)),3)  = read_INWD_mo_acc_child_care
				inwd_array((Ubound(inwd_array,1)),4)  = read_INWD_mo_acc_med_support
				inwd_array((Ubound(inwd_array,1)),5)  = read_INWD_mo_acc_othr_support
				inwd_array((Ubound(inwd_array,1)),6)  = read_INWD_nonmo_acc_basic_support
				inwd_array((Ubound(inwd_array,1)),7)  = read_INWD_nonmo_acc_spousal_maint
				inwd_array((Ubound(inwd_array,1)),8)  = read_INWD_nonmo_acc_child_care
				inwd_array((Ubound(inwd_array,1)),9)  = read_INWD_nonmo_acc_med_support
				inwd_array((Ubound(inwd_array,1)),10) = read_INWD_nonmo_acc_othr_support
				inwd_array((Ubound(inwd_array,1)),11) = read_INWD_additional_20pct
				inwd_array((Ubound(inwd_array,1)),12) = read_INWD_ttl_iw_amt
					
					'Advance the employer array counter one position
				employer_array_position = employer_array_position + 1
				
				CALL go_to("B", 3, 29)
			END IF
			row = row + 1
		LOOP UNTIL row = 20 OR end_of_data = "*** End of Data ***"
		PF8
		row = 7
	LOOP UNTIL end_of_data = "*** End of Data ***"

END FUNCTION

'----INWD----
FUNCTION read_INWD_employer_name
	EMReadScreen employer_name, 30, 8, 7
	employer_name = trim(employer_name)
	If employer_name <> "" Then read_INWD_employer_name = employer_name
	IF employer_name =  "" Then read_INWD_employer_name = "EMPLOYER NOT FOUND"
END FUNCTION

FUNCTION read_INWD_mo_acc_basic_support
	EMReadScreen moacc_basic_support, 14, 15, 17
	moacc_basic_support = trim(moacc_basic_support)
	IF moacc_basic_support <> "" THEN read_INWD_mo_acc_basic_support = moacc_basic_support
	IF moacc_basic_support =  "" THEN read_INWD_mo_acc_basic_support = "0.00"
END FUNCTION

FUNCTION read_INWD_mo_acc_spousal_maint
	EMReadScreen moacc_spou_main, 14, 16, 17
	moacc_spou_main = trim(moacc_spou_main)
	IF moacc_spou_main <> "" THEN read_INWD_mo_acc_spousal_maint = moacc_spou_main
	IF moacc_spou_main =  "" THEN read_INWD_mo_acc_spousal_maint = "0.00"
END FUNCTION

FUNCTION read_INWD_mo_acc_child_care
	EMReadScreen moacc_child_care, 14, 17, 17
	moacc_child_care = trim(moacc_child_care)
	IF moacc_child_care <> "" THEN read_INWD_mo_acc_child_care = moacc_child_care
	IF moacc_child_care =  "" THEN read_INWD_mo_acc_child_care = "0.00"
END FUNCTION

FUNCTION read_INWD_mo_acc_med_support
	EMReadScreen moacc_med_support, 14, 18, 17
	moacc_med_support = trim(moacc_med_support)
	IF moacc_med_support <> "" THEN read_INWD_mo_acc_med_support = moacc_med_support
	IF moacc_med_support =  "" THEN read_INWD_mo_acc_med_support = "0.00"
END FUNCTION

FUNCTION read_INWD_mo_acc_othr_support
	EMReadScreen moacc_othr_support, 14, 19, 17
	moacc_othr_support = trim(moacc_othr_support)
	IF moacc_othr_support <> "" THEN read_INWD_mo_acc_othr_support = moacc_othr_support
	IF moacc_othr_support =  "" THEN read_INWD_mo_acc_othr_support = "0.00"
END FUNCTION

FUNCTION read_INWD_nonmo_acc_basic_support
	EMReadScreen nonmoacc_basic_support, 15, 15, 32
	nonmoacc_basic_support = trim(nonmoacc_basic_support)
	IF nonmoacc_basic_support <> "" THEN read_INWD_nonmo_acc_basic_support = nonmoacc_basic_support
	IF nonmoacc_basic_support =  "" THEN read_INWD_nonmo_acc_basic_support = "0.00"
END FUNCTION

FUNCTION read_INWD_nonmo_acc_spousal_maint
	EMReadScreen nonmoacc_spou_main, 15, 16, 32
	nonmoacc_spou_main = trim(nonmoacc_spou_main)
	IF nonmoacc_spou_main <> "" THEN read_INWD_nonmo_acc_spousal_maint = nonmoacc_spou_main
	IF nonmoacc_spou_main =  "" THEN read_INWD_nonmo_acc_spousal_maint = "0.00"
END FUNCTION

FUNCTION read_INWD_nonmo_acc_child_care
	EMReadScreen nonmoacc_child_care, 15, 17, 32
	nonmoacc_child_care = trim(nonmoacc_child_care)
	IF nonmoacc_child_care <> "" THEN read_INWD_nonmo_acc_child_care = nonmoacc_child_care
	IF nonmoacc_child_care =  "" THEN read_INWD_nonmo_acc_child_care = "0.00"
END FUNCTION

FUNCTION read_INWD_nonmo_acc_med_support
	EMReadScreen nonmoacc_med_support, 15, 18, 32
	nonmoacc_med_support = trim(nonmoacc_med_support)
	IF nonmoacc_med_support <> "" THEN read_INWD_nonmo_acc_med_support = nonmoacc_med_support
	IF nonmoacc_med_support =  "" THEN read_INWD_nonmo_acc_med_support = "0.00"
END FUNCTION

FUNCTION read_INWD_nonmo_acc_othr_support
	EMReadScreen nonmoacc_othr_support, 15, 19, 32
	nonmoacc_othr_support = trim(nonmoacc_othr_support)
	IF nonmoacc_othr_support <> "" THEN read_INWD_nonmo_acc_othr_support = nonmoacc_othr_support
	IF nonmoacc_othr_support =  "" THEN read_INWD_nonmo_acc_othr_support = "0.00"
END FUNCTION

FUNCTION read_INWD_additional_20pct
	EMReadScreen add_20pct, 14, 19, 48
	add_20pct = trim(add_20pct)
	IF add_20pct <> "" THEN read_INWD_additional_20pct = add_20pct
	IF add_20pct =  "" THEN read_INWD_additional_20pct = "0.00"
END FUNCTION

FUNCTION read_INWD_ttl_iw_amt
	EMReadScreen ttl_iw_amt, 15, 19, 64
	ttl_iw_amt = trim(ttl_iw_amt)
	IF ttl_iw_amt <> "" THEN read_INWD_ttl_iw_amt = ttl_iw_amt
	IF ttl_iw_amt =  "" THEN read_INWD_ttl_iw_amt = "0.00"
END FUNCTION

FUNCTION go_to(value, row, col)
	EMWriteScreen value, row, col
	transmit
END FUNCTION

FUNCTION build_dialog(employer_number,inwd_array)
	'Using the array and the position number of the employer this sets the dialog values to the correct values
BeginDialog inwd_dialog, 0, 0, 196, 300, "INWD Dialog"
  ButtonGroup ButtonPressed
    PushButton 5, 5, 30, 15, "CAFS", CAFS_button
    PushButton 95, 5, 30, 15, "NCDD", NCDD_button
    PushButton 125, 5, 30, 15, "PALC", PALC_button
    PushButton 35, 5, 30, 15, "ENFL", ENFL_button
    PushButton 65, 5, 30, 15, "INWD", INWD_button
    CancelButton 160, 5, 30, 15
  Text 10, 30, 40, 10, "Employer:"
  Text 75, 30, 85, 10, inwd_array(employer_number, 0)
  Text 10, 50, 75, 10, "Monthly Accrual"
  Text 20, 65, 50, 10, "Basic Support"
  Text 110, 65, 30, 10, inwd_array(employer_number, 1)
  Text 20, 80, 60, 10, "Spousal Maint."
  Text 110, 80, 30, 10, inwd_array(employer_number, 2)
  Text 20, 95, 60, 10, "Child Care"
  Text 110, 95, 30, 10, inwd_array(employer_number, 3)
  Text 20, 110, 60, 10, "Medical Support"
  Text 110, 110, 30, 10, inwd_array(employer_number, 4)
  Text 20, 125, 60, 10, "Other Support"
  Text 110, 125, 30, 10, inwd_array(employer_number, 5)
  Text 10, 145, 75, 10, "Monthly Accrual"
  Text 20, 160, 50, 10, "Basic Support"
  Text 110, 160, 30, 10, inwd_array(employer_number, 6)
  Text 20, 175, 60, 10, "Spousal Support"
  Text 110, 175, 30, 10, inwd_array(employer_number, 7)
  Text 20, 190, 60, 10, "Child Care"
  Text 110, 190, 30, 10, inwd_array(employer_number, 8)
  Text 20, 205, 60, 10, "Medical Support"
  Text 110, 205, 30, 10, inwd_array(employer_number, 9)
  Text 20, 220, 60, 10, "Other Support"
  Text 110, 220, 30, 10, inwd_array(employer_number, 10)
  Text 10, 240, 75, 10, "Additional 20%"
  Text 110, 240, 30, 10, inwd_array(employer_number, 11)
  Text 10, 255, 75, 10, "Total IW Amount"
  Text 110, 255, 30, 10, inwd_array(employer_number, 12)
EndDialog
	


	DIALOG inwd_dialog
END FUNCTION

'----ENFL----
FUNCTION read_ENFL_case_based_remedy(case_based_remedy)
	'Going to ENFL
	CALL navigate_to_PRISM_screen("ENFL")
	CALL go_to("Y", 20, 74)
	'Gathering the case-based remedies
	row = 8
	DO
		msgbox row
		EMReadScreen end_of_data, 11, row, 32
		IF end_of_data <> "End of Data" THEN 
			EMReadScreen ENFL_case_number, 12, row, 67
			ENFL_case_number = replace(ENFL_case_number, " ", "")
			EMReadScreen remedy, 3, row, 2
			remedy = replace(remedy, " ", "")
			IF ENFL_case_number = replace(PRISM_case_number, " ", "") THEN case_based_remedy = case_based_remedy & remedy & ", "
		END IF
		row = row + 1
		IF row = 20 THEN
			PF8
			row = 8
		END IF
	LOOP UNTIL end_of_data = "End of Data"
END FUNCTION

FUNCTION read_ENFL_person_based_remedy(person_based_remedy)
	CALL navigate_to_PRISM_screen("ENFL")
	CALL go_to("Y", 20, 74)
	row = 8
	DO
		msgbox row
		EMReadScreen end_of_data, 11, row, 32
		IF end_of_data <> "End of Data" THEN 
			EMReadScreen ENFL_case_number, 12, row, 67
			ENFL_case_number = trim(ENFL_case_number)
			IF ENFL_case_number = "" THEN 
				EMReadScreen remedy, 3, row, 2
				remedy = trim(remedy)
				IF remedy <> "" THEN person_based_remedy = person_based_remedy & remedy & ", "
			END IF
		END IF
		row = row + 1
		IF row = 20 THEN 
			PF8
			row = 8
		END IF
	LOOP UNTIL end_of_data = "End of Data"
END FUNCTION

'----The functions that pull it all together----
FUNCTION create_NCDD_variable(NCDD)
	CALL read_NCDD_address_known(addr_known)
	IF addr_known = "Y" THEN
		CALL read_NCDD_address_effective_date(addr_date)
		NCDD = "Address known, last verified: " & addr_date
	ELSE
		NCDD = "Address not known."
	END IF
END FUNCTION


FUNCTION create_CAFS_variable(CAFS)
	CALL navigate_to_PRISM_screen("CAFS")
	CALL read_monthly_accrual(monthly_accrual)	'For whatever reason, it has to read monthly_accrual twice...not sure why

	CALL read_monthly_accrual(monthly_accrual)
	cafs = cafs & "Mo. Acc: " & monthly_accrual & "; "
	CALL read_monthly_non_accrual(mo_non_acc)
	cafs = cafs & "Mo. Non-Acc: " & mo_non_acc & "; "
	CALL read_total_arrears(total_arrears)
	cafs = cafs & "Ttl Arrears: " & total_arrears & "; "
	CALL read_npa_arrears(npa_arrears)
	cafs = cafs & "NPA Arrears: " & npa_arrears & "; "
	CALL read_pa_arrears(pa_arrears)
	cafs = cafs & "PA Arrears: " & pa_arrears
END FUNCTION

FUNCTION create_PALC_variable(PALC)
	CALL navigate_to_PRISM_screen("PALC")
	CALL read_PALC_last_payment_date(last_payment_date)
	PALC = PALC & "Last Payment Date: " & last_payment_date & "; "
	CALL read_PALC_payment_type(payment_type)
	PALC = PALC & "Payment Type: " & payment_type & "; "
	CALL read_PALC_payment_amount(payment_amount)
	PALC = PALC & "Payment Amount: " & payment_amount & "; "
	CALL read_PALC_alloc_amount(alloc_amount)
	PALC = PALC & "Case Allocated Amount: " & alloc_amount
END FUNCTION

FUNCTION create_ENFL_variable(ENFL)
	CALL read_ENFL_case_based_remedy(ENFL_case_based)
	IF ENFL_case_based = "" THEN 
		ENFL = "Case Based: None;"
	ELSE
		ENFL = "Case Based: " & ENFL_case_based & "; "
	END IF
	CALL read_ENFL_person_based_remedy(ENFL_person_based)
	IF ENFL_person_based = "" THEN 
		ENFL = ENFL & "Person Based: None"
	ELSE
		ENFL = ENFL & "Person Based: " & person_based
	END IF
END FUNCTION

'-----dialogs-----
FUNCTION all_dialogs(dialog_name)
	IF dialog_name = "CASE NUMBER" THEN
		EditBox 115, 20, 80, 15, PRISM_case_number
		BeginDialog dialog_name, 0, 0, 201, 65, "CASE NUMBER"
		ButtonGroup ButtonPressed
			OkButton 50, 45, 50, 15
			CancelButton 100, 45, 50, 15
		Text 10, 10, 95, 10, "No case number was found."
		Text 10, 25, 105, 10, "Please enter the case number:"
		EndDialog
	ELSEIF dialog_name = "MENU" THEN
		BeginDialog dialog_name, 0, 0, 296, 215, "MENU"
		ButtonGroup ButtonPressed
			PushButton 15, 20, 30, 15, "CAFS", CAFS_button
			PushButton 45, 20, 30, 15, "ENFL", ENFL_button
			PushButton 75, 20, 30, 15, "INWD", INWD_button
			PushButton 105, 20, 30, 15, "LETL", LETL_button
			PushButton 135, 20, 30, 15, "NCDD", NCDD_button
			PushButton 165, 20, 30, 15, "NCID", NCID_button
			PushButton 195, 20, 30, 15, "PALC", PALC_button
			PushButton 225, 20, 30, 15, "PAPD", PAPD_button
			PushButton 255, 20, 30, 15, "SUDL", SUDL_button
			CancelButton 260, 195, 30, 15
		Text 10, 60, 280, 10, "This is the Payment Review script. You can use it to quickly review information on:"
		Text 15, 75, 25, 10, "* CAFS"
		Text 45, 75, 25, 10, "* ENFL"
		Text 75, 75, 25, 10, "* INWD"
		Text 105, 75, 25, 10, "* LETL"
		Text 135, 75, 25, 10, "*NCDD"
		Text 165, 75, 25, 10, "* NCID"
		Text 195, 75, 25, 10, "* PALC"
		Text 225, 75, 25, 10, "* PAPD"
		Text 255, 75, 25, 10, "* SUDL"
		Text 10, 95, 275, 20, "After reviewing the information, you will be prompted to write what the next action on the case will be."
		Text 10, 120, 220, 10, "You can also send specific DORD documents."
		Text 10, 140, 220, 10, "If, at any point, you wish to stop the script, press ``Cancel``"
		Text 10, 160, 220, 10, "Use the DISPLAY BUTTONS to begin."
		GroupBox 10, 10, 280, 35, "Display Buttons"
		EndDialog
	ELSEIF dialog_name = "CAFS" THEN
BeginDialog dialog_name, 0, 0, 296, 355, "CAFS"
  ButtonGroup ButtonPressed
    PushButton 15, 15, 30, 15, "CAFS", CAFS_nav_button
    PushButton 45, 15, 30, 15, "ENFL", ENFL_nav_button
    PushButton 75, 15, 30, 15, "INWD", INWD_nav_button
    PushButton 105, 15, 30, 15, "LETL", LETL_nav_button
    PushButton 135, 15, 30, 15, "NCDD", NCDD_nav_button
    PushButton 165, 15, 30, 15, "NCID", NCID_nav_button
    PushButton 195, 15, 30, 15, "PALC", PALC_nav_button
    PushButton 225, 15, 30, 15, "PAPD", PAPD_nav_button
    PushButton 255, 15, 30, 15, "SUDL", SUDL_nav_button
    PushButton 15, 60, 30, 15, "CAFS", CAFS_button
    PushButton 45, 60, 30, 15, "ENFL", ENFL_button
    PushButton 75, 60, 30, 15, "INWD", INWD_button
    PushButton 105, 60, 30, 15, "LETL", LETL_button
    PushButton 135, 60, 30, 15, "NCDD", NCDD_button
    PushButton 165, 60, 30, 15, "NCID", NCID_button
    PushButton 195, 60, 30, 15, "PALC", PALC_button
    PushButton 225, 60, 30, 15, "PAPD", PAPD_button
    PushButton 255, 60, 30, 15, "SUDL", SUDL_button
    OkButton 230, 335, 30, 15
    CancelButton 260, 335, 30, 15
  GroupBox 10, 50, 280, 35, "Display Buttons"
  GroupBox 10, 5, 280, 35, "Navigation Buttons"
  GroupBox 150, 230, 40, 60, "Extra NAV"
    GroupBox 10, 225, 130, 45, "DORD Docs"
  Text 65, 105, 60, 10, "Monthly Accrual"
  Text 170, 105, 65, 10, monthly_accrual
  Text 65, 120, 75, 10, "Monthly Non-Accrual"
  Text 170, 120, 65, 10, mo_non_acc
  Text 65, 135, 75, 10, "Total Arrears"
  Text 170, 135, 65, 10, total_arrears
  Text 170, 150, 65, 10, npa_arrears
  Text 170, 165, 65, 10, pa_arrears
  Text 65, 150, 75, 10, "NPA Arrears"
  Text 65, 165, 75, 10, "PA Arrears"
  Text 15, 205, 35, 10, "Results: "
  EditBox 55, 200, 235, 15, review_result
  CheckBox 15, 235, 120, 10, "Send Non-Compliance with DLPP", non_compliance_check
  CheckBox 15, 245, 120, 10, "Send Address Verification", addr_verif_check
  CheckBox 15, 255, 120, 10, "Send Non-Pay", non_pay_check
    PushButton 155, 245, 30, 10, "CAAD", CAAD_nav_button
    PushButton 155, 255, 30, 10, "CAHL", CAHL_nav_button
    PushButton 155, 265, 30, 10, "MAXIS",MAXIS_nav_button
    PushButton 155, 275, 30, 10, "MMIS", MMIS_nav_button
EndDialog
	ELSEIF dialog_name = "ENFL" THEN
BeginDialog dialog_name, 0, 0, 296, 355, "ENFL"
  ButtonGroup ButtonPressed
    PushButton 15, 15, 30, 15, "CAFS", CAFS_nav_button
    PushButton 45, 15, 30, 15, "ENFL", ENFL_nav_button
    PushButton 75, 15, 30, 15, "INWD", INWD_nav_button
    PushButton 105, 15, 30, 15, "LETL", LETL_nav_button
    PushButton 135, 15, 30, 15, "NCDD", NCDD_nav_button
    PushButton 165, 15, 30, 15, "NCID", NCID_nav_button
    PushButton 195, 15, 30, 15, "PALC", PALC_nav_button
    PushButton 225, 15, 30, 15, "PAPD", PAPD_nav_button
    PushButton 255, 15, 30, 15, "SUDL", SUDL_nav_button
    PushButton 15, 60, 30, 15, "CAFS", CAFS_button
    PushButton 45, 60, 30, 15, "ENFL", ENFL_button
    PushButton 75, 60, 30, 15, "INWD", INWD_button
    PushButton 105, 60, 30, 15, "LETL", LETL_button
    PushButton 135, 60, 30, 15, "NCDD", NCDD_button
    PushButton 165, 60, 30, 15, "NCID", NCID_button
    PushButton 195, 60, 30, 15, "PALC", PALC_button
    PushButton 225, 60, 30, 15, "PAPD", PAPD_button
    PushButton 255, 60, 30, 15, "SUDL", SUDL_button
    OkButton 230, 335, 30, 15
    CancelButton 260, 335, 30, 15
  GroupBox 10, 50, 280, 35, "Display Buttons"
  GroupBox 10, 5, 280, 35, "Navigation Buttons"
  GroupBox 150, 230, 40, 60, "Extra NAV"
  GroupBox 10, 225, 130, 45, "DORD Docs"
  Text 155, 105, 95, 10, ENFL_case_based
  Text 155, 120, 95, 10, ENFL_person_based
  Text 15, 205, 35, 10, "Results: "
  EditBox 55, 200, 235, 15, review_result
  CheckBox 15, 235, 120, 10, "Send Non-Compliance with DLPP", non_compliance_check
  CheckBox 15, 245, 120, 10, "Send Address Verification", addr_verif_check
  CheckBox 15, 255, 120, 10, "Send Non-Pay", non_pay_check
    PushButton 155, 245, 30, 10, "CAAD", CAAD_nav_button
    PushButton 155, 255, 30, 10, "CAHL", CAHL_nav_button
    PushButton 155, 265, 30, 10, "MAXIS", MAXIS_nav_button
    PushButton 155, 275, 30, 10, "MMIS", MMIS_nav_button
  Text 25, 105, 100, 10, "Case-Based Remedies:"
  Text 25, 120, 100, 10, "Person-Based Remedies:"
EndDialog


	ELSEIF dialog_name = "NCDD EDIT" THEN
		BeginDialog dialog_name, 0, 0, 296, 250, "NCDD (Edit Mode)"
		ButtonGroup ButtonPressed
			PushButton 15, 15, 30, 15, "CAFS", CAFS_nav_button
			PushButton 45, 15, 30, 15, "ENFL", ENFL_nav_button
			PushButton 75, 15, 30, 15, "INWD", INWD_nav_button
			PushButton 105, 15, 30, 15, "LETL", LETL_nav_button
			PushButton 135, 15, 30, 15, "NCDD", NCDD_nav_button
			PushButton 165, 15, 30, 15, "NCID", NCID_nav_button
			PushButton 195, 15, 30, 15, "PALC", PALC_nav_button
			PushButton 225, 15, 30, 15, "PAPD", PAPD_nav_button
			PushButton 255, 15, 30, 15, "SUDL", SUDL_nav_button
			CancelButton 260, 205, 30, 15
		Text 10, 35, 60, 10, "Address Known?"
		EditBox 115, 35, 65, 15, addr_known
		Text 10, 50, 75, 10, "Date Last Verified"
		EditBox 115, 50, 65, 15, addr_date
			PushButton 215, 50, 70, 15, "FINISHED EDITING", done_ncdd_edit_button
		EndDialog
	ELSEIF dialog_name = "NCDD" THEN
		BeginDialog dialog_name, 0, 0, 296, 355, "NCDD"
		ButtonGroup ButtonPressed
			PushButton 15, 15, 30, 15, "CAFS", CAFS_nav_button
			PushButton 45, 15, 30, 15, "ENFL", ENFL_nav_button
			PushButton 75, 15, 30, 15, "INWD", INWD_nav_button
			PushButton 105, 15, 30, 15, "LETL", LETL_nav_button
			PushButton 135, 15, 30, 15, "NCDD", NCDD_nav_button
			PushButton 165, 15, 30, 15, "NCID", NCID_nav_button
			PushButton 195, 15, 30, 15, "PALC", PALC_nav_button
			PushButton 225, 15, 30, 15, "PAPD", PAPD_nav_button
			PushButton 255, 15, 30, 15, "SUDL", SUDL_nav_button
			PushButton 15, 60, 30, 15, "CAFS", CAFS_button
			PushButton 45, 60, 30, 15, "ENFL", ENFL_button
			PushButton 75, 60, 30, 15, "INWD", INWD_button
			PushButton 105, 60, 30, 15, "LETL", LETL_button
			PushButton 135, 60, 30, 15, "NCDD", NCDD_button
			PushButton 165, 60, 30, 15, "NCID", NCID_button
			PushButton 195, 60, 30, 15, "PALC", PALC_button
			PushButton 225, 60, 30, 15, "PAPD", PAPD_button
			PushButton 255, 60, 30, 15, "SUDL", SUDL_button
			PushButton 250, 100, 30, 15, "EDIT", ncdd_edit_button
			OkButton 230, 235, 30, 15
			CancelButton 260, 235, 30, 15
		Text 40, 105, 60, 10, "Address Known?"
		Text 145, 105, 65, 10, addr_known
		Text 40, 120, 75, 10, "Date Last Verified"
		Text 145, 120, 65, 10, addr_date
		GroupBox 10, 50, 280, 35, "Display Buttons"
		GroupBox 10, 5, 280, 35, "Navigation Buttons"
		GroupBox 150, 230, 40, 60, "Extra NAV"
		GroupBox 10, 225, 130, 45, "DORD Docs"
		Text 15, 205, 35, 10, "Results: "
		EditBox 55, 200, 235, 15, review_result
		CheckBox 15, 235, 120, 10, "Send Non-Compliance with DLPP", non_compliance_check
		CheckBox 15, 245, 120, 10, "Send Address Verification", addr_verif_check
		CheckBox 15, 255, 120, 10, "Send Non-Pay", non_pay_check
		PushButton 155, 245, 30, 10, "CAAD", CAAD_nav_button
		PushButton 155, 255, 30, 10, "CAHL", CAHL_nav_button
		PushButton 155, 265, 30, 10, "MAXIS",MAXIS_nav_button
		PushButton 155, 275, 30, 10, "MMIS", MMIS_nav_button
		EndDialog
		
	ELSEIF dialog_name = "SUDL" THEN
		SUDL_display = split(SUDL, ";")
		SUDL_dlg_row = 125
	
		BeginDialog dialog_name, 0, 0, 296, 260, "SUDL"
			ButtonGroup ButtonPressed
				PushButton 15, 15, 30, 15, "CAFS", CAFS_nav_button
				PushButton 45, 15, 30, 15, "ENFL", ENFL_nav_button
				PushButton 75, 15, 30, 15, "INWD", INWD_nav_button
				PushButton 105, 15, 30, 15, "LETL", LETL_nav_button
				PushButton 135, 15, 30, 15, "NCDD", NCDD_nav_button
				PushButton 165, 15, 30, 15, "NCID", NCID_nav_button
				PushButton 195, 15, 30, 15, "PALC", PALC_nav_button
				PushButton 225, 15, 30, 15, "PAPD", PAPD_nav_button
				PushButton 255, 15, 30, 15, "SUDL", SUDL_nav_button
				PushButton 15, 60, 30, 15, "CAFS", CAFS_button
				PushButton 45, 60, 30, 15, "ENFL", ENFL_button
				PushButton 75, 60, 30, 15, "INWD", INWD_button
				PushButton 105, 60, 30, 15, "LETL", LETL_button
				PushButton 135, 60, 30, 15, "NCDD", NCDD_button
				PushButton 165, 60, 30, 15, "NCID", NCID_button
				PushButton 195, 60, 30, 15, "PALC", PALC_button
				PushButton 225, 60, 30, 15, "PAPD", PAPD_button
				PushButton 255, 60, 30, 15, "SUDL", SUDL_button
				OkButton 230, 235, 30, 15
				CancelButton 260, 235, 30, 15
			GroupBox 10, 50, 280, 35, "Display Buttons"
			GroupBox 10, 5, 280, 35, "Navigation Buttons"
			Text 10, 105, 125, 10, "Suppressed Enforcement Remedies:"
				
			FOR EACH enf_rem IN SUDL_display
				Text 40, SUDL_dlg_row, 150, 10, enf_rem		
				SUDL_dlg_row = SUDL_dlg_row + 15
			NEXT
		EndDialog
	ELSEIF dialog_name = "LETL" THEN 
		LETL_display = split(LETL, ";")
		LETL_dlg_row = 125
	
		BeginDialog dialog_name, 0, 0, 296, 260, "LETL"
			ButtonGroup ButtonPressed
				PushButton 15, 15, 30, 15, "CAFS", CAFS_nav_button
				PushButton 45, 15, 30, 15, "ENFL", ENFL_nav_button
				PushButton 75, 15, 30, 15, "INWD", INWD_nav_button
				PushButton 105, 15, 30, 15, "LETL", LETL_nav_button
				PushButton 135, 15, 30, 15, "NCDD", NCDD_nav_button
				PushButton 165, 15, 30, 15, "NCID", NCID_nav_button
				PushButton 195, 15, 30, 15, "PALC", PALC_nav_button
				PushButton 225, 15, 30, 15, "PAPD", PAPD_nav_button
				PushButton 255, 15, 30, 15, "SUDL", SUDL_nav_button
				PushButton 15, 60, 30, 15, "CAFS", CAFS_button
				PushButton 45, 60, 30, 15, "ENFL", ENFL_button
				PushButton 75, 60, 30, 15, "INWD", INWD_button
				PushButton 105, 60, 30, 15, "LETL", LETL_button
				PushButton 135, 60, 30, 15, "NCDD", NCDD_button
				PushButton 165, 60, 30, 15, "NCID", NCID_button
				PushButton 195, 60, 30, 15, "PALC", PALC_button
				PushButton 225, 60, 30, 15, "PAPD", PAPD_button
				PushButton 255, 60, 30, 15, "SUDL", SUDL_button
				OkButton 230, 235, 30, 15
				CancelButton 260, 235, 30, 15
			GroupBox 10, 50, 280, 35, "Display Buttons"
			GroupBox 10, 5, 280, 35, "Navigation Buttons"

			Text 10, 105, 125, 10, "Legal Tracking List:"
			
			FOR EACH letl_line IN LETL_display
				Text 40, LETL_dlg_row, 150, 10, letl_line
				LETL_dlg_row = LETL_dlg_row + 15
			NEXT
		EndDialog
	END IF

	DIALOG dialog_name
END FUNCTION

FUNCTION all_buttons(current_dlg)
	IF ButtonPressed = CAFS_button THEN 
		CALL all_dialogs("CAFS")
		current_dlg = "CAFS"
		page_count = page_count + 1
	ELSEIF ButtonPressed = LETL_button THEN
		CALL all_dialogs("LETL")
		current_dlg = "LETL"
		page_count = page_count + 1
	ELSEIF ButtonPressed = NCDD_button THEN 
		CALL all_dialogs("NCDD")
		current_dlg = "NCDD"
		page_count = page_count + 1
	ELSEIF ButtonPressed = SUDL_button THEN 
		CALL all_dialogs("SUDL")
		current_dlg = "SUDL"
		page_count = page_count + 1
	ELSEIF ButtonPressed = ENFL_button THEN
		CALL all_dialogs("ENFL")
		current_dlg = "ENFL"
		page_count = page_count + 1
	ELSEIF ButtonPressed = CAFS_nav_button THEN 
		CALL navigate_to_PRISM_screen("CAFS")
		CALL all_dialogs(current_dlg)
		page_count = page_count + 1
	ELSEIF ButtonPressed = ENFL_nav_button THEN 
		CALL navigate_to_PRISM_screen("ENFL")
		CALL all_dialogs(current_dlg)
		page_count = page_count + 1
	ELSEIF ButtonPressed = INWD_nav_button THEN 
		CALL navigate_to_PRISM_screen("INWD")
		CALL all_dialogs(current_dlg)
		page_count = page_count + 1
	ELSEIF ButtonPressed = LETL_nav_button THEN 
		CALL navigate_to_PRISM_screen("LETL")
		CALL all_dialogs(current_dlg)
		page_count = page_count + 1
	ELSEIF ButtonPressed = NCDD_nav_button THEN 
		CALL navigate_to_PRISM_screen("NCDD")
		CALL all_dialogs(current_dlg)
		page_count = page_count + 1
	ELSEIF ButtonPressed = NCID_nav_button THEN 
		CALL navigate_to_PRISM_screen("NCID")
		CALL all_dialogs(current_dlg)
		page_count = page_count + 1
	ELSEIF ButtonPressed = PALC_nav_button THEN 
		CALL navigate_to_PRISM_screen("PALC")
		CALL all_dialogs(current_dlg)
		page_count = page_count + 1
	ELSEIF ButtonPressed = PAPD_nav_button THEN 
		CALL navigate_to_PRISM_screen("PAPD")
		CALL all_dialogs(current_dlg)
		page_count = page_count + 1
	ELSEIF ButtonPressed = SUDL_nav_button THEN 
		CALL navigate_to_PRISM_screen("SUDL")
		CALL all_dialogs(current_dlg)
		page_count = page_count + 1
	ELSEIF ButtonPressed = CAAD_nav_button THEN
		CALL navigate_to_PRISM_screen("CAAD")
		CALL all_dialogs(current_dlg)
		page_count = page_count + 1
	ELSEIF ButtonPressed = CAHL_nav_button THEN
		CALL navigate_to_PRISM_screen("CAHL")
		CALL all_dialogs(current_dlg)
		page_count = page_count + 1
'	ELSEIF ButtonPressed = MAXIS_nav_button THEN
'		Here's where we need a custom function that pulls the NCP's SSN and navigates to MAXIS Inquiry DB to search PERS by SSN and then ...
'			for all case numbers with "AP" search CASE/CURR for status <> "INACTIVE" and then search CASE/PERS for cases where status <> "INACTIVE" for the cl with that SSN active on CASH.
'		CALL all_dialogs(current_dlg)
'	ELSEIF ButtonPressed = MMIS_nav_button THEN
'		Here's where we need a custom function that pulls the NCP's SSN and navigates to MMIS to search for an active Medicaid or MinnesotaCare case.
'		CALL all_dialogs(current_dlg)
	END IF
	

END FUNCTION

'-------------------------------------------------------------------------------THE SCRIPT---------------------------------------------------------------------------------
EMConnect ""

CALL navigate_to_PRISM_screen("CAAD")
CALL find_variable("Case: ", PRISM_case_number, 13)
DO
	CALL all_dialogs("CASE NUMBER")
		IF ButtonPressed = 0 THEN stopscript
LOOP UNTIL PRISM_case_number <> ""
CALL go_to(PRISM_case_number, 20, 8)

CALL create_CAFS_variable(CAFS)
CALL create_ENFL_variable(ENFL)
ReDim inwd_array(0, 12)
CALL create_INWD_array
CALL create_LETL_variable(LETL, 8)
CALL create_NCDD_variable(NCDD)
CALL create_PALC_variable(PALC)
CALL create_SUDL_variable(SUDL, 8)

msgbox lenENFL_case_based
msgbox ENFL_person_based

page_count = 0
DO
	IF page_count = 0 THEN CALL all_dialogs("MENU")
	IF ButtonPressed = 0 THEN 
		cancel_warning = MsgBox("Are you sure you want to cancel? Press YES to cancel. Press NO to return to the script.", vbYesNo)
		IF cancel_warning = vbYes THEN stopscript
	END IF
	IF ButtonPressed <> 0 AND ButtonPressed <> -1 THEN CALL all_buttons(current_dlg)
LOOP UNTIL ButtonPressed = -1

'Outputs one dialog box per employer as defined by the number of employers
FOR i = 0 TO (UBound(inwd_array,1))
		'I = the array position of the current employer and passes this to the dialog box with the full array
		CALL build_dialog(i,inwd_array)
NEXT

msgbox "done"
stopscript
