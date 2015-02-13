DIM total_arrears
DIM npa_arrears
DIM pa_arrears
DIM mo_non_acc
DIM monthly_accrual
DIM addr_date
DIM addr_known

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
	'Getting the case number -- until a better option is available.
	CALL navigate_to_PRISM_screen("CAAD")
	CALL find_variable("Case: ", PRISM_case_number, 13)
	PRISM_case_number = replace(PRISM_case_number, " ", "")

	'Going back to ENFL
	CALL navigate_to_PRISM_screen("ENFL")
	CALL go_to("Y", 20, 74)
	'Gathering the case-based remedies
	row = 8
	DO
		DO
			EMReadScreen end_of_data, 11, 24, 2
			EMReadScreen ENFL_case_number, 12, row, 67
			ENFL_case_number = replace(ENFL_case_number, " ", "")
			EMReadScreen remedy, 3, row, 2
			remedy = replace(remedy, " ", "")
			IF ENFL_case_number = PRISM_case_number THEN case_based_remedy = case_based_remedy & remedy & ", "
			row = row + 1
		LOOP UNTIL row = 20
		PF8
		row = 8
	LOOP UNTIL end_of_data = "End of data"
END FUNCTION

FUNCTION read_ENFL_person_based_remedy(person_based_remedy)
	CALL navigate_to_PRISM_screen("ENFL")
	CALL go_to("Y", 20, 74)
	row = 8
	DO
		DO
			EMReadScreen end_of_data, 11, 24, 2
			EMReadScreen ENFL_case_number, 12, row, 67
			ENFL_case_number = trim(ENFL_case_number)
			IF ENFL_case_number = "" THEN 
				EMReadScreen remedy, 3, row, 2
				remedy = trim(remedy)
				IF remedy <> "" THEN person_based_remedy = person_based_remedy & remedy & ", "
			END IF
			row = row + 1
		LOOP UNTIL row = 20
		PF8
		row = 8
	LOOP UNTIL end_of_data = "End of data"
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
	CALL read_ENFL_case_based_remedy(case_based)
	IF case_based <> "" THEN ENFL = "Case Based: " & case_based & "; "
	CALL read_ENFL_person_based_remedy(person_based)
	IF person_based <> "" THEN ENFL = ENFL & "Person Based: " & person_based
END FUNCTION

'-----dialogs-----
FUNCTION all_dialogs(dialog_name)
	IF dialog_name = "CAFS" THEN
		BeginDialog dialog_name, 0, 0, 196, 185, "CAFS"
		  ButtonGroup ButtonPressed
		    PushButton 5, 5, 30, 15, "CAFS", CAFS_button
		    PushButton 95, 5, 30, 15, "NCDD", NCDD_button
		    PushButton 125, 5, 30, 15, "PALC", PALC_button
		    PushButton 35, 5, 30, 15, "ENFL", ENFL_button
		    PushButton 65, 5, 30, 15, "INWD", INWD_button
		    CancelButton 160, 5, 30, 15
		  Text 10, 35, 60, 10, "Monthly Accrual"
		  Text 115, 35, 65, 10, monthly_accrual
		  Text 10, 50, 75, 10, "Monthly Non-Accrual"
		  Text 115, 50, 65, 10, mo_non_acc
		  Text 10, 65, 75, 10, "Total Arrears"
		  Text 115, 65, 65, 10, total_arrears
		  Text 115, 80, 65, 10, npa_arrears
		  Text 115, 95, 65, 10, pa_arrears
		  Text 10, 80, 75, 10, "NPA Arrears"
		  Text 10, 95, 75, 10, "PA Arrears"
		EndDialog
	ELSEIF dialog_name = "MENU" THEN
		BeginDialog dialog_name, 0, 0, 196, 185, "MENU"
		  ButtonGroup ButtonPressed
		    PushButton 5, 5, 30, 15, "CAFS", CAFS_button
		    PushButton 95, 5, 30, 15, "NCDD", NCDD_button
		    PushButton 125, 5, 30, 15, "PALC", PALC_button
		    PushButton 35, 5, 30, 15, "ENFL", ENFL_button
		    PushButton 65, 5, 30, 15, "INWD", INWD_button
		    CancelButton 160, 5, 30, 15
		EndDialog
	ELSEIF dialog_name = "NCDD EDIT" THEN
		BeginDialog dialog_name, 0, 0, 196, 185, "NCDD"
		  ButtonGroup ButtonPressed
		    PushButton 5, 5, 30, 15, "CAFS", CAFS_button
		    PushButton 95, 5, 30, 15, "NCDD", NCDD_button
		    PushButton 125, 5, 30, 15, "PALC", PALC_button
		    PushButton 35, 5, 30, 15, "ENFL", ENFL_button
		    PushButton 65, 5, 30, 15, "INWD", INWD_button
		    CancelButton 160, 5, 30, 15
		  Text 10, 35, 60, 10, "Address Known?"
		  EditBox 115, 35, 65, 15, addr_known
		  Text 10, 50, 75, 10, "Date Last Verified"
		  EditBox 115, 50, 65, 15, addr_date
		    PushButton 135, 165, 60, 15, "FINISHED EDITING", done_ncdd_edit_button
		EndDialog
	ELSEIF dialog_name = "NCDD" THEN
		BeginDialog dialog_name, 0, 0, 196, 185, "NCDD"
		  ButtonGroup ButtonPressed
		    PushButton 5, 5, 30, 15, "CAFS", CAFS_button
		    PushButton 95, 5, 30, 15, "NCDD", NCDD_button
		    PushButton 125, 5, 30, 15, "PALC", PALC_button
		    PushButton 35, 5, 30, 15, "ENFL", ENFL_button
		    PushButton 65, 5, 30, 15, "INWD", INWD_button
		    CancelButton 160, 5, 30, 15
		  Text 10, 35, 60, 10, "Address Known?"
		  Text 115, 35, 65, 10, addr_known
		  Text 10, 50, 75, 10, "Date Last Verified"
		  Text 115, 50, 65, 10, addr_date
		    PushButton 155, 165, 30, 15, "EDIT", ncdd_edit_button
		EndDialog
	END IF

	DIALOG dialog_name
END FUNCTION

BeginDialog Dialog1, 0, 0, 336, 150, "Dialog"
  ButtonGroup ButtonPressed
    PushButton 5, 10, 30, 15, "CAFS", CAFS_button
  EditBox 50, 10, 275, 15, CAFS
  ButtonGroup ButtonPressed
    PushButton 5, 30, 30, 15, "NCDD", NCDD_button
  EditBox 50, 30, 275, 15, NCDD
  ButtonGroup ButtonPressed
    PushButton 5, 50, 30, 15, "PALC", PALC_button
  EditBox 50, 50, 275, 15, PALC
  ButtonGroup ButtonPressed
    PushButton 5, 70, 30, 15, "ENFL", ENFL_button
  EditBox 50, 70, 275, 15, ENFL
  ButtonGroup ButtonPressed
    PushButton 115, 130, 50, 15, "INWD", INWD_button
    CancelButton 170, 130, 50, 15
EndDialog



EMConnect "B"
CALL navigate_to_PRISM_screen("CAFS")
CALL create_CAFS_variable(CAFS)
CALL create_PALC_variable(PALC)
CALL create_NCDD_variable(NCDD)
CALL create_ENFL_variable(ENFL)
ReDim inwd_array(0, 12)
CALL create_INWD_array


DO
	CALL all_dialogs("MENU")
	IF ButtonPressed = CAFS_button THEN CALL all_dialogs("CAFS")
	IF ButtonPressed = PALC_button THEN CALL navigate_to_PRISM_screen("PALC")
	IF ButtonPressed = NCDD_button THEN CALL all_dialogs("NCDD")
	IF ButtonPressed = ncdd_edit_button THEN CALL all_dialogs("NCDD EDIT")
	IF ButtonPressed = done_ncdd_edit_button THEN call all_dialogs("NCDD")
	IF ButtonPressed = ENFL_button THEN CALL navigate_to_PRISM_screen("ENFL")
	IF ButtonPressed = 0 THEN stopscript
LOOP UNTIL ButtonPressed = INWD_button

'Outputs one dialog box per employer as defined by the number of employers
FOR i = 0 TO (UBound(inwd_array,1))
		'I = the array position of the current employer and passes this to the dialog box with the full array
		CALL build_dialog(i,inwd_array)
NEXT
