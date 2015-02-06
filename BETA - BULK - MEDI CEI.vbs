'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "BULK - MEDI CEI"
start_time = timer

'FUNCTIONS----------------------------------------------------------------------------------------------------
'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
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
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

BeginDialog CEI_dialog, 0, 0, 216, 400, "CEI dialog"
  EditBox 15, 20, 70, 15, case01
  EditBox 15, 40, 70, 15, case02
  EditBox 15, 60, 70, 15, case03
  EditBox 15, 80, 70, 15, case04
  EditBox 15, 100, 70, 15, case05
  EditBox 15, 120, 70, 15, case06
  EditBox 15, 140, 70, 15, case07
  EditBox 15, 160, 70, 15, case08
  EditBox 15, 180, 70, 15, case09
  EditBox 15, 200, 70, 15, case10
  EditBox 15, 220, 70, 15, case11
  EditBox 15, 240, 70, 15, case12
  EditBox 15, 260, 70, 15, case13
  EditBox 15, 280, 70, 15, case14
  EditBox 15, 300, 70, 15, case15
  EditBox 100, 20, 15, 15, memb01
  EditBox 150, 20, 45, 15, medi_amt01
  EditBox 100, 40, 15, 15, memb02
  EditBox 150, 40, 45, 15, medi_amt02
  EditBox 100, 60, 15, 15, memb03
  EditBox 150, 60, 45, 15, medi_amt03
  EditBox 100, 80, 15, 15, memb04
  EditBox 150, 80, 45, 15, medi_amt04
  EditBox 100, 100, 15, 15, memb05
  EditBox 150, 100, 45, 15, medi_amt05
  EditBox 100, 120, 15, 15, memb06
  EditBox 150, 120, 45, 15, medi_amt06
  EditBox 100, 140, 15, 15, memb07
  EditBox 150, 140, 45, 15, medi_amt07
  EditBox 100, 160, 15, 15, memb08
  EditBox 150, 160, 45, 15, medi_amt08
  EditBox 100, 180, 15, 15, memb09
  EditBox 150, 180, 45, 15, medi_amt09
  EditBox 100, 200, 15, 15, memb10
  EditBox 150, 200, 45, 15, medi_amt10
  EditBox 100, 220, 15, 15, memb11
  EditBox 150, 220, 45, 15, medi_amt11
  EditBox 100, 240, 15, 15, memb12
  EditBox 150, 240, 45, 15, medi_amt12
  EditBox 100, 260, 15, 15, memb13
  EditBox 150, 260, 45, 15, medi_amt13
  EditBox 100, 280, 15, 15, memb14
  EditBox 150, 280, 45, 15, medi_amt14
  EditBox 100, 300, 15, 15, memb15
  EditBox 150, 300, 45, 15, medi_amt15
  ButtonGroup ButtonPressed
    PushButton 75, 320, 120, 15, "Autofill Memb 01 and 104.90 for all", autofill_memb_medi
  EditBox 140, 340, 50, 15, reimbursement_month
  EditBox 140, 360, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 65, 380, 50, 15
    CancelButton 120, 380, 50, 15
  Text 15, 5, 50, 10, "Case number: "
  Text 100, 5, 30, 10, "Memb #:"
  Text 150, 5, 60, 10, "Medicare amount:"
  Text 25, 345, 110, 10, "Reimbursement month (MM/YY):"
  Text 70, 365, 65, 10, "Sign the case note:"
EndDialog

'----------THE SCRIPT----------
EMConnect ""

maxis_check_function
back_to_self

DO
	DO
		Dialog CEI_dialog
			IF buttonpressed = 0 then stopscript
			IF worker_signature = "" THEN MsgBox "Please sign your case notes."
			IF reimbursement_month = "" THEN MsgBox "Please enter a month to check for reimbursement."
			IF len(reimbursement_month) <> 5 THEN MsgBox "Please enter a reimbursement month in the format MM/YY."
			IF ButtonPressed = autofill_memb_medi THEN
				IF case01 <> "" THEN
					medi_amt01 = "104.90"
					memb01 = "01"
				END IF
				IF case02 <> "" THEN 
					medi_amt02 = "104.90"
					memb02 = "01"
				END IF
				IF case03 <> "" THEN 
					medi_amt03 = "104.90" 
					memb03 = "01"
				END IF
				IF case04 <> "" THEN
					medi_amt04 = "104.90"
					memb04 = "01"
				END IF
				IF case05 <> "" THEN 
					medi_amt05 = "104.90" 
					memb05 = "01"
				END IF
				IF case06 <> "" THEN 
					medi_amt06 = "104.90" 
					memb06 = "01"
				END IF
				IF case07 <> "" THEN 
					medi_amt07 = "104.90" 
					memb07 = "01"
				END IF
				IF case08 <> "" THEN 
					medi_amt08 = "104.90"
					memb08 = "01"
				END IF
				IF case09 <> "" THEN 
					medi_amt09 = "104.90"
					memb09 = "01"
				END IF
				IF case10 <> "" THEN 
					medi_amt10 = "104.90"
					memb10 = "01"
				END IF
				IF case11 <> "" THEN 
					medi_amt11 = "104.90"
					memb11 = "01"
				END IF
				IF case12 <> "" THEN 
					medi_amt12 = "104.90"
					memb12 = "01"
				END IF
				IF case13 <> "" THEN 
					medi_amt13 = "104.90"
					memb13 = "01"
				END IF
				IF case14 <> "" THEN 
					medi_amt14 = "104.90"
					memb14 = "01"
				END IF
				IF case15 <> "" THEN 
					medi_amt15 = "104.90"
					memb15 = "01"
				END IF
			END IF
	LOOP UNTIL (worker_signature <> "") AND (reimbursement_month <> "") AND (len(reimbursement_month) = 5)
LOOP UNTIL ButtonPressed = -1

bene_month = left(reimbursement_month, 2)
bene_year = right(reimbursement_month, 2)

IF case01 <> "" THEN 
	IF len(case01) <> 8 THEN
		DO
			case01 = "0" & case01
		LOOP UNTIL len(case01) = 8
	END IF
	cei_array = cei_array & memb01 & case01 & medi_amt01 & "~"
END IF
IF case02 <> "" THEN
	IF len(case02) <> 8 THEN
		DO
			case02 = "0" & case02
		LOOP UNTIL len(case02) = 8
	END IF
	cei_array = cei_array & memb02 & case02 & medi_amt02 & "~"
END IF
IF case03 <> "" THEN
	IF len(case03) <> 8 THEN
		DO
			case03 = "0" & case03
		LOOP UNTIL len(case03) = 8
	END IF
	cei_array = cei_array & memb03 & case03 & medi_amt03 & "~"
END IF
IF case04 <> "" THEN 
	IF len(case04) <> 8 THEN
		DO
			case04 = "0" & case04
		LOOP UNTIL len(case04) = 8
	END IF
	cei_array = cei_array & memb04 & case04 & medi_amt04 & "~"
END IF
IF case05 <> "" THEN 
	IF len(case05) <> 8 THEN
		DO
			case05 = "0" & case05
		LOOP UNTIL len(case05) = 8
	END IF
	cei_array = cei_array & memb05 & case05 & medi_amt05 & "~"
END IF
IF case06 <> "" THEN 
	IF len(case06) <> 8 THEN
		DO
			case06 = "0" & case06
		LOOP UNTIL len(case06) = 8
	END IF
	cei_array = cei_array & memb06 & case06 & medi_amt06 & "~"
END IF
IF case07 <> "" THEN 
	IF len(case07) <> 8 THEN
		DO
			case07 = "0" & case07
		LOOP UNTIL len(case07) = 8
	END IF
	cei_array = cei_array & memb07 & case07 & medi_amt07 & "~"
END IF
IF case08 <> "" THEN 
	IF len(case08) <> 8 THEN
		DO
			case08 = "0" & case08
		LOOP UNTIL len(case08) = 8
	END IF
	cei_array = cei_array & memb08 & case08 & medi_amt08 & "~"
END IF
IF case09 <> "" THEN 
	IF len(case09) <> 8 THEN
		DO
			case09 = "0" & case09
		LOOP UNTIL len(case09) = 8
	END IF
	cei_array = cei_array & memb09 & case09 & medi_amt09 & "~"
END IF
IF case10 <> "" THEN 
	IF len(case10) <> 8 THEN
		DO
			case10 = "0" & case10
		LOOP UNTIL len(case10) = 8
	END IF
	cei_array = cei_array & memb10 & case10 & medi_amt10 & "~"
END IF
IF case11 <> "" THEN 
	IF len(case11) <> 8 THEN
		DO
			case11 = "0" & case11
		LOOP UNTIL len(case11) = 8
	END IF
	cei_array = cei_array & memb11 & case11 & medi_amt11 & "~"
END IF
IF case12 <> "" THEN
	IF len(case12) <> 8 THEN
		DO
			case12 = "0" & case12
		LOOP UNTIL len(case12) = 8
	END IF
	cei_array = cei_array & memb12 & case12 & medi_amt12 & "~"
END IF
IF case13 <> "" THEN 
	IF len(case13) <> 8 THEN
		DO
			case13 = "0" & case13
		LOOP UNTIL len(case13) = 8
	END IF
	cei_array = cei_array & memb13 & case13 & medi_amt13 & "~"
END IF
IF case14 <> "" THEN 
	IF len(case14) <> 8 THEN
		DO
			case14 = "0" & case14
		LOOP UNTIL len(case14) = 8
	END IF
	cei_array = cei_array & memb14 & case14 & medi_amt14 & "~"
END IF
IF case15 <> "" THEN 
	IF len(case15) <> 8 THEN
		DO
			case15 = "0" & case15
		LOOP UNTIL len(case15) = 8
	END IF
	cei_array = cei_array & memb15 & case15 & medi_amt15 & "~"
END IF

cei_array = split(cei_array, "~")


'Now it checks to make sure MAXIS production is running on this screen.
attn
Do
	EMReadScreen MAI_check, 3, 1, 33
	If MAI_check <> "MAI" then EMWaitReady 1, 1
Loop until MAI_check = "MAI"

EMReadScreen production_check, 7, 6, 15

IF production_check <> "RUNNING" THEN script_end_procedure("You do not appear to be running production version of MAXIS.") 

EMReadScreen mmis_a_check, 7, 15, 15
IF mmis_a_check = "RUNNING" THEN
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
EMFocus	'Bringing the window focus to the second screen if needed

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

FOR EACH cei_case IN cei_array
 IF cei_case <> "" THEN
	memb_num = left(cei_case, 2)
	case_num = right(left(cei_case, 10), 8)
	EMWriteScreen "i", 2, 19
	EMWriteScreen case_num, 9, 19
	transmit			'navigates to RCAD
	transmit			'navigates to RREP
	transmit			'navigates to RCIN

	RCIN_row = 11
	DO
		EMReadScreen rel_num, 2, RCIN_row, 13
		IF rel_num <> memb_num THEN 
			RCIN_row = RCIN_row + 1
			IF RCIN_row = 22 THEN failure_array = failure_array & ("Case " & case_num & " --- Member #" & memb_num & " not found on MMIS/RCIN.") & "~~~"
		END IF
	LOOP UNTIL (rel_num = memb_num) OR (RCIN_row = 22)
	IF rel_num = memb_num THEN
		EMWriteScreen "X", RCIN_row, 2
		transmit		'navigates to RSUM
		EMWriteScreen "RELG", 1, 8
		transmit		'navigates to RELG

		EMReadScreen prog01_type, 8, 6, 13
		EMReadScreen prog02_type, 8, 10, 13
		EMReadScreen prog03_type, 8, 14, 13
		EMReadScreen prog04_type, 8, 18, 13

		IF prog01_type = "MEDICAID" THEN
			EMReadScreen elig01_type, 2, 6, 33
			EMReadScreen elig01_end, 8, 7, 36
			IF elig01_type = "DP" AND elig01_end <> "99/99/99" THEN failure_array = failure_array & ("Case " & case_num & " --- Member #" & memb_num & " has a closed MA-EPD span on MMIS/RELG.") & "~~~"
		ELSE
			IF prog02_type = "MEDICAID" THEN
				EMReadScreen elig02_type, 2, 10, 33
				EMReadScreen elig02_end, 8, 11, 36
				IF elig02_type = "DP" AND elig02_end <> "99/99/99" THEN failure_array = failure_array & ("Case " & case_num & " --- Member #" & memb_num & " has a closed MA-EPD span on MMIS/RELG.") & "~~~"
			ELSE
				IF prog03_type = "MEDICAID" THEN
					EMReadScreen elig03_type, 2, 14, 33
					EMReadScreen elig03_end, 8, 15, 36
					IF elig03_type = "DP" AND elig03_end <> "99/99/99" THEN failure_array = failure_array & ("Case " & case_num & " --- Member #" & memb_num & " has a closed MA-EPD span on MMIS/RELG.") & "~~~"
				ELSE
					IF prog04_type = "MEDICAID" THEN
						EMReadScreen elig04_type, 2, 18, 33
						EMReadScreen elig04_end, 8, 19, 36
						IF elig04_type = "DP" AND elig04_end <> "99/99/99" THEN failure_array = failure_array & ("Case " & case_num & " --- Member #" & memb_num & " has a closed MA-EPD span on MMIS/RELG.") & "~~~"
					ELSE
						failure_array = failure_array & ("Case " & case_num & " --- Member #" & memb_num & " is not open on Medicaid on MMIS/RELG.") & "~~~"
					END IF
				END IF
			END IF
		END IF
		IF (prog01_type = "MEDICAID" OR prog02_type = "MEDICAID" OR prog03_type = "MEDICAID" OR prog04_type = "MEDICAID") AND (elig01_type <> "DP" AND elig02_type <> "DP" AND elig03_type <> "DP" AND elig04_type <> "DP") THEN failure_array = failure_array & ("Case " & case_num & " --- Member #" & memb_num & " is open on MA but NOT MA-EPD on MMIS/RELG.") & "~~~"
	
		'Condition for success to move to RBYB
		IF (prog01_type = "MEDICAID" AND elig01_type = "DP" AND elig01_end = "99/99/99") OR (prog02_type = "MEDICAID" AND elig02_type = "DP" AND elig02_end = "99/99/99") OR (prog03_type = "MEDICAID" AND elig03_type = "DP" AND elig03_end = "99'99'99") OR (prog04_type = "MEDICAID" AND elig04_type = "DP" AND elig04_end = "99/99/99") THEN 

			EMWriteScreen "RBYB", 1, 8
			transmit		'navigates to RBYB to check if the CL is already open on the buy-in.

			EMReadScreen accrete_date, 8, 5, 66
			EMReadScreen delete_date, 8, 6, 65
			accrete_date = replace(accrete_date, " ", "")
			IF accrete_date <> "" & delete_date = "99/99/99" THEN failure_array = failure_array & ("Case " & case_num & " --- Member #" & memb_num & " is still on the Buy-In on MMIS/RBYB.") & "~~~"
			IF (accrete_date = "") OR (accrete_date <> "" & delete_date <> "99/99/99") THEN 		'Condition for success to move to RMSC
				EMWriteScreen "RMSC", 1, 8
				transmit		'navigates to RMSC to check

				EMReadScreen ssi_begin_dt, 5, 5, 28
					ssi_begin_dt = replace(ssi_begin_dt, " ", "")
				EMReadScreen ssi_end_dt, 5, 5, 66
					ssi_begin_dt = replace(ssi_end_dt, " ", "")
				EMReadScreen msa_grh_begin_dt, 5, 8, 28
					msa_grh_begin_dt = replace(msa_grh_begin_dt, " ", "")
				EMReadScreen msa_grh_end_dt, 5, 8, 66
					msa_grh_end_dt = replace(msa_grh_end_dt, " ", "")

				IF ssi_begin_dt <> "" AND ssi_end_dt <> "99/99" THEN failure_array = failure_array & ("Case " & case_num & " --- Member #" & memb_num & " has an SSI Begin Date but no SSI End Date on MMIS/RMSC.") & "~~~"
				IF msa_grh_begin_dt <> "" AND msa_grh_end_dt <> "99/99" THEN failure_array = failure_array & ("Case " & case_num & " --- Member #" & memb_num & " has an MSA/GRH Begin Date but no MSA/GRH End Date on MMIS/RMSC.") & "~~~"

				IF (ssi_begin_dt = "" OR (ssi_begin_dt <> "" AND ssi_end_dt = "99/99")) AND (msa_grh_begin_dt = "" OR (msa_grh_begin_dt <> "" AND msa_grh_end_dt = "99/99")) THEN success_array = success_array & cei_case & "~~~"
			END IF
		END IF
	END IF
	DO
		EMReadScreen mmis_recip_key, 14, 1, 31
		IF mmis_recip_key <> "MMIS RECIP KEY" THEN PF3
	LOOP UNTIL mmis_recip_key = "MMIS RECIP KEY"
 END IF
NEXT

success_array = split(success_array, "~~~")

'----------Navigates back to MAXIS----------
attn
Do
	EMReadScreen MAI_check, 3, 1, 33
	If MAI_check <> "MAI" then EMWaitReady 1, 1
Loop until MAI_check = "MAI"
EMReadScreen maxis_prod_check, 7, 6, 15
IF maxis_prod_check = "RUNNING" THEN
	EMWriteScreen "1", 2, 15
	transmit
ELSE
	EMConnect"A"
	attn
	Do
		EMReadScreen MAI_check, 3, 1, 33
		If MAI_check <> "MAI" then EMWaitReady 1, 1
	Loop until MAI_check = "MAI"
	EMReadScreen maxis_prod_check, 7, 6, 15
	IF maxis_prod_check = "RUNNING" THEN
		EMWriteScreen "1", 2, 15
		transmit
	ELSE
		EMConnect"B"
		attn
		Do
			EMReadScreen MAI_check, 3, 1, 33
			If MAI_check <> "MAI" then EMWaitReady 1, 1
		Loop until MAI_check = "MAI"
		EMReadScreen maxis_prod_check, 7, 6, 15
		IF maxis_prod_check = "RUNNING" THEN
			EMWriteScreen "1", 2, 15
			transmit
		END IF
	END IF
END IF

Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.add()
Set objSelection = objWord.Selection

'----------Goes through MAXIS to determine if the CL is open on MA-EPD in MAXIS and that they are eligible for reimbursement----------
FOR EACH good_case IN success_array
 IF good_case <> "" THEN
	memb_num = left(good_case, 2)
	case_number = right(left(good_case, 10), 8)
	medi_amt = right(good_case, (len(good_case) - 10))		
	
	maxis_check_function
	back_to_SELF
	
	call navigate_to_screen("ELIG", "HC")
	EMWriteScreen bene_month, 20, 56
	EMWriteScreen bene_year, 20, 59
	transmit

	hhmm_row = 8
	DO
		EMReadScreen hhmm_memb, 2, hhmm_row, 3
		IF hhmm_memb <> memb_num THEN	hhmm_row = hhmm_row + 1
	LOOP UNTIL hhmm_row = 20 OR hhmm_memb = memb_num

	IF hhmm_memb = memb_num THEN
		EMWriteScreen "X", hhmm_row, 26
		transmit
		bsum_col = 19
		DO
			EMReadScreen bsum_month, 5, 6, bsum_col
			IF bsum_month <> reimbursement_month THEN
				bsum_col = bsum_col + 11
				IF (bsum_col = 85) AND (bsum_month <> reimbursement_month) THEN failure_array = failure_array & ("Case " & case_number & " --- Member #" & memb_num & " does not have HC results in " & reimbursement_month & " in BSUM.") & "~~~"
			END IF
		LOOP UNTIL (bsum_month = reimbursement_month) OR (bsum_col = 85)

		EMReadScreen bsum_elig_type, 2, 12, (bsum_col - 2)
		IF bsum_elig_type <> "DP" THEN failure_array = failure_array & ("Case " & case_number & " --- Member #" & memb_num & " is not open on MA-EPD in BSUM.") & "~~~"
		IF bsum_elig_type = "DP" THEN
			EMWriteScreen "X", 9, (bsum_col + 2)	'Navigates to the budget for the reimbursement month
			transmit
		
			EMReadScreen pct_fpg, 3, 18, 39
			pct_fpg = pct_fpg * 1
			IF pct_fpg > 200 THEN failure_array = failure_array & ("Case " & case_number & " --- Member #" & memb_num & " is over 200% FPG on EBUD.") & "~~~"
			IF pct_fpg <= 200 THEN
'				objSelection.TypeText ("This would be the case note in case " & case_number)
'				objSelection.Paragraph()
'				objSelection.TypeText ("Medicare reimbursement for " & reimbursement_month & " sent to fiscal")				'commenting for development mode
'				objSelection.TypeParagraph()
'				objSelection.TypeText ("* Medicare amount: " & medi_amt)
'				objSelection.TypeParagraph()
'				objSelection.TypeText ("---")
'				objSelection.TypeParagraph()
'				objSelection.TypeText (worker_signature)
'				objSelection.TypeParagraph()
'				objSelection.TypeParagraph()
'				objSelection.TypeParagraph()
'				objSelection.TypeParagraph()
				PF4
				PF9
				call write_new_line_in_case_note("Medicare reimbursement for " & reimbursement_month & " sent to fiscal")
				call write_new_line_in_case_note("* Medicare amount: " & medi_amt)
				call write_new_line_in_case_note("---")
				call write_new_line_in_case_note(worker_signature)
				'PF3	
			END IF		
		END IF
	ELSE
		failure_array = failure_array & ("Case " & case_number & " --- Member #" & memb_num & " does not have HC open in ELIG/HC.") & "~~~"
	END IF
 END IF
NEXT

failure_array = trim(failure_array)
failure_array = split(failure_array, "~~~")

'----------Now the script creates the Word Document is used to report on the cases that fail the checks----------
Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.add()
Set objSelection = objWord.Selection

objSelection.typetext "Errors were found in the following cases..."
objSelection.TypeParagraph()
objSelection.TypeParagraph()

FOR EACH failed_case IN failure_array
	objSelection.TypeText failed_case
	objSelection.TypeParagraph()
NEXT

MsgBox "Your cases have been case noted! Don't forget to send the authorization for payment to fiscal."

script_end_procedure("")
