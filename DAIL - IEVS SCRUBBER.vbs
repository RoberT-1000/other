'Stats========================
name_of_script = "DAIL - IEVS SCRUBBER.vbs"
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

FUNCTION income_matrix(income_matrix_array, ievs_name, ievs_employer, quarter, quarterly_wage, ievs_year, all_programs)
	CALL navigate_to_MAXIS_screen("STAT", "MEMB")
	DO
		EMReadScreen ref_num, 2, 4, 33
		ref_num = trim(ref_num)
		IF ref_num <> "" THEN client_array = client_array & ref_num & "|"
		transmit
		EMReadScreen error_msg, 10, 24, 2
		error_msg = trim(error_msg)
	LOOP UNTIL error_msg <> ""
	
	client_array = trim(client_array)
	client_array = split(client_array, "|")
	
	'income_matrix_array positions:
	'		(x, 0) = client name and reference number
	'		(x, 1) = panel number
	'		(x, 2) = employer
	'		(x, 3) = retro
	'		(x, 4) = prosp
	'		(x, 5) = pic
	'		(x, 6) = hc inc est
	'		(x, 8) = quarterly retro total
	'		(x, 9) = quarterly prosp total
	'		(x, 10) = quarterly pic
	'		(x, 11) = quarterly hc inc est
	
	num_of_income_panels = 0
	ReDim income_matrix_array(30, 6)
	
	IF quarter = 1 THEN 
		months_array = "01" & "/" & right(ievs_year, 2) & ",02" & "/" & right(ievs_year, 2) & ",03" & "/" & right(ievs_year, 2) & ""
	ELSEIF quarter = 2 THEN 
		months_array = "04" & "/" & right(ievs_year, 2) & ",05" & "/" & right(ievs_year, 2) & ",06" & "/" & right(ievs_year, 2) & ""
	ELSEIF quarter = 3 THEN
		months_array = "07" & "/" & right(ievs_year, 2) & ",08" & "/" & right(ievs_year, 2) & ",09" & "/" & right(ievs_year, 2) & ""
	ELSEIF quarter = 4 THEN 
		months_array = "10" & "/" & right(ievs_year, 2) & ",11" & "/" & right(ievs_year, 2) & ",12" & "/" & right(ievs_year, 2) & ""
	END IF
	
	'adding CM and CM+1 to array
	current_month = "0" & datepart("M", date)
	current_month = right(current_month, 2)
	current_month = current_month & "/" & right(datepart("YYYY", date), 2)
	current_month_plus_one = "0" & datepart("M", dateadd("M", 1, date))
	current_month_plus_one = right(current_month_plus_one, 2)
	current_month_plus_one = current_month_plus_one & "/" & right(datepart("YYYY", dateadd("M", 1, date)), 2)
	months_array = months_array & "," & current_month & "," & current_month_plus_one
		
	months_array = split(months_array, ",")
	
	'JOBS
	FOR EACH ievs_month IN months_array
		back_to_SELF
		EMWriteScreen left(ievs_month, 2), 20, 43
		EMWriteScreen right(ievs_month, 2), 20, 46
		CALL navigate_to_MAXIS_screen("STAT", "JOBS")
		'Checking to make sure the client was active in that particular month
		EMReadScreen invalid_month, 60, 24, 2
		IF InStr(invalid_month, "INVALID") = 0 THEN 
			EMWriteScreen "01", 20, 76
			EMWriteScreen "01", 20, 79
			transmit
			FOR EACH ref_num IN client_array
				IF ref_num <> "" THEN 
					EMWriteScreen ref_num, 20, 76
					EMWriteScreen "01", 20, 79
					transmit
					EMReadScreen cl_name, 30, 4, 36
					cl_name = trim(cl_name)
					IF cl_name = ievs_name THEN 
						EMReadScreen num_of_jobs, 1, 2, 78
						IF num_of_jobs <> "0" THEN 
							DO
								EMReadScreen jobs_end_date, 8, 9, 49
								jobs_end_date = replace(jobs_end_date, " ", "/")
								first_of_month = left(ievs_month, 2) & "/01/" & right(ievs_month, 2)
								IF jobs_end_date = "__/__/__" THEN  
									num_of_income_panels = num_of_income_panels + 1
									income_matrix_array(num_of_income_panels, 0) = ref_num & " " & cl_name
									EMReadScreen jobs_num, 2, 2, 72
									jobs_num = replace(jobs_num, " ", "0")
									income_matrix_array(num_of_income_panels, 1) = "JOBS " & jobs_num
									EMReadScreen employer, 30, 7, 42
									employer = replace(employer, "_", "")
									income_matrix_array(num_of_income_panels, 2) = employer & " (" & ievs_month & ")"
									EMReadScreen retro_amt, 8, 17, 38
									retro_amt = trim(retro_amt)
									IF retro_amt = "" THEN retro_amt = 0.00
									income_matrix_array(num_of_income_panels, 3) = retro_amt
									EMReadScreen prosp_amt, 8, 17, 67
									prosp_amt = trim(prosp_amt)
									IF prosp_amt = "" THEN prosp_amt = 0.00
									income_matrix_array(num_of_income_panels, 4) = prosp_amt
									'Getting in to the PIC
									EMWriteScreen "X", 19, 38
									transmit
									EMReadScreen pic_amt, 8, 18, 56
									pic_amt = trim(pic_amt)
									IF pic_amt = "" THEN pic_amt = 0.00
									income_matrix_array(num_of_income_panels, 5) = pic_amt
									PF3
									
									'>>>>> GRABBING THE HC BUDGET INFO <<<<<
									'Grabbing the pay frequency
									EMReadScreen pay_freq, 1, 18, 35
									'Going into the HC Inc Est
									EMWriteScreen "X", 19, 54
									transmit
									'Reading the budgetted amount
									EMReadScreen hc_inc_est, 8, 11, 63
									hc_inc_est = replace(hc_inc_est, "_", "")
									hc_inc_est = trim(hc_inc_est)
									IF hc_inc_est = "" THEN hc_inc_est = 0.00
									'Converting the budgetted amount per pay period to the monthly amount
									IF pay_freq = "1" THEN 
										income_matrix_array(num_of_income_panels, 6) = hc_inc_est
									ELSEIF pay_freq = "2" OR pay_freq = "3" THEN 
										income_matrix_array(num_of_income_panels, 6) = 2 * hc_inc_est
									ELSEIF pay_freq = "4" THEN 
										income_matrix_array(num_of_income_panels, 6) = 4 * hc_inc_est
									END IF
									'Exiting the HC inc est	
									transmit
									
								ELSEIF (jobs_end_date <> "__/__/__" AND DateDiff("D", jobs_end_date, first_of_month) < 0) THEN
									num_of_income_panels = num_of_income_panels + 1
									income_matrix_array(num_of_income_panels, 0) = ref_num & " " & cl_name
									EMReadScreen jobs_num, 2, 2, 72
									jobs_num = replace(jobs_num, " ", "0")
									income_matrix_array(num_of_income_panels, 1) = "JOBS " & jobs_num
									EMReadScreen employer, 30, 7, 42
									employer = replace(employer, "_", "")
									income_matrix_array(num_of_income_panels, 2) = employer & " (" & ievs_month & ")"
									EMReadScreen retro_amt, 8, 17, 38
									retro_amt = trim(retro_amt)
									IF retro_amt = "" THEN retro_amt = 0.00
									income_matrix_array(num_of_income_panels, 3) = retro_amt
									EMReadScreen prosp_amt, 8, 17, 67
									prosp_amt = trim(prosp_amt)
									IF prosp_amt = "" THEN prosp_amt = 0.00
									income_matrix_array(num_of_income_panels, 4) = prosp_amt
									'Getting in to the PIC
									EMWriteScreen "X", 19, 38
									transmit
									EMReadScreen pic_amt, 8, 18, 56
									pic_amt = trim(pic_amt)
									IF pic_amt = "" THEN pic_amt = 0.00
									income_matrix_array(num_of_income_panels, 5) = pic_amt
									PF3
									
									'>>>>> GRABBING THE HC BUDGET INFO <<<<<
									'Grabbing the pay frequency
									EMReadScreen pay_freq, 1, 18, 35
									'Going into the HC Inc Est
									EMWriteScreen "X", 19, 54
									transmit
									'Reading the budgetted amount
									EMReadScreen hc_inc_est, 8, 11, 63
									hc_inc_est = replace(hc_inc_est, "_", "")
									hc_inc_est = trim(hc_inc_est)
									IF hc_inc_est = "" THEN hc_inc_est = 0.00
									'Converting the budgetted amount per pay period to the monthly amount
									IF pay_freq = "1" THEN 
										income_matrix_array(num_of_income_panels, 6) = hc_inc_est
									ELSEIF pay_freq = "2" OR pay_freq = "3" THEN 
										income_matrix_array(num_of_income_panels, 6) = 2 * hc_inc_est
									ELSEIF pay_freq = "4" THEN 
										income_matrix_array(num_of_income_panels, 6) = 4 * hc_inc_est
									END IF
									'Exiting the HC inc est	
									transmit
								END IF
								'Going to the next JOBS panel
								transmit
								EMReadScreen error_msg, 10, 24, 2
								error_msg = trim(error_msg)
							LOOP UNTIL error_msg <> ""		
						END IF
					END IF
				END IF
			NEXT
		END IF
	NEXT
	
	' >>>>> DETERMINING THE NUMBER OF UNIQUE EMPLOYERS FOR THE OUTPUT <<<<<
	all_employers_array = ""
	FOR i = 1 TO num_of_income_panels
		current_employer = left(income_matrix_array(i, 2), len(income_matrix_array(i, 2)) - 7)
		IF InStr(all_employers_array, current_employer) = 0 THEN all_employers_array = all_employers_array & current_employer & ","
	NEXT
	
	all_employers_array = all_employers_array & "~~"
	all_employers_array = replace(all_employers_array, ",~~", "")
	all_employers_array = trim(all_employers_array)
	all_employers_array = split(all_employers_array, ",")
	
	number_of_unique_employers = ubound(all_employers_array)
	
	REDIM quarterly_output_array(number_of_unique_employers, 4)
	'contents of each position for this array
	' (x, 0) = employer
	' (x, 1) = quarterly retro
	' (x, 2) = quarterly prosp
	' (x, 3) = quarterly PIC
	' (x, 4) = quarterly HC
	
	'Adding the employer names to the output array
	employer_position = 0
	FOR EACH unique_employer IN all_employers_array
		quarterly_output_array(employer_position, 0) = unique_employer
		employer_position = employer_position + 1
	NEXT
	
	'Adding quarterly wage information to the output array
	FOR unique_job = 0 TO number_of_unique_employers
		'Setting the positions to numeric values
		quarterly_output_array(unique_job, 1) = 0
		quarterly_output_array(unique_job, 2) = 0
		quarterly_output_array(unique_job, 3) = 0
		quarterly_output_array(unique_job, 4) = 0
	NEXT
	
	FOR unique_job = 0 TO number_of_unique_employers		
		'Creating quarterly totals for each unique jaeorb
		FOR i = 1 TO num_of_income_panels
			IF (current_month & ")") <> right(income_matrix_array(i, 2), 6) AND (current_month_plus_one & ")") <> right(income_matrix_array(i, 2), 6) AND _
				TRIM(quarterly_output_array(unique_job, 0)) = TRIM(left(income_matrix_array(i, 2), len(income_matrix_array(i, 2)) - 7)) THEN   'this will make sure we only add the first 3 months (the ones from the quarter in question)
					quarterly_output_array(unique_job, 1) = quarterly_output_array(unique_job, 1) + income_matrix_array(i, 3)		'Quarterly Retro for that job
					quarterly_output_array(unique_job, 2) = quarterly_output_array(unique_job, 2) + income_matrix_array(i, 4)		'Quarterly Prosp for that job
					quarterly_output_array(unique_job, 3) = quarterly_output_array(unique_job, 3) + income_matrix_array(i, 5)		'Quarterly PIC for that job
					quarterly_output_array(unique_job, 4) = quarterly_output_array(unique_job, 4) + income_matrix_array(i, 6)		'Quarterly HC for that job
			END IF
		NEXT
	NEXT
	
	dlg_height = 150 + (num_of_income_panels * 15)
			
	BeginDialog income_matrix_dlg, 0, 0, 510, dlg_height, "Income Matrix"
		ButtonGroup ButtonPressed
			OkButton 350, (135 + (15 * num_of_income_panels)), 50, 15
			CancelButton 400, (135 + (15 * num_of_income_panels)), 50, 15
		Text 15, 15, 60, 10, "IEVS Client"
		Text 15, 25, 120, 10, ievs_name
		Text 150, 15, 60, 10, "IEVS Employer"
		Text 150, 25, 160, 10, ievs_employer
		Text 330, 15, 30, 10, "Quarter"
		Text 345, 25, 20, 10, "Q" & quarter
		Text 380, 15, 60, 10, "Quarterly Wages"
		Text 385, 25, 30, 10, quarterly_wage
		Text 450, 15, 60, 10, "Programs"
		Text 450, 25, 60, 10, all_programs
		Text 15, 45, 80, 10, "Quarterly Earnings, Q" & quarter
		Text 150, 45, 120, 10, "Sum of Q" & quarter & " Earnings"
		Text 330, 45, 30, 10, "Retro"
		Text 370, 45, 30, 10, "Prosp"
		Text 410, 45, 30, 10, "PIC"
		Text 450, 45, 40, 10, "Quarterly HC"
		'Displaying quarterly wage info for each jaeorb
		FOR z = 0 TO number_of_unique_employers
			Text 15, 55 + (10 * z), 120, 10, ievs_name
			Text 150, 55 + (10 * z), 160, 10, quarterly_output_array(z, 0)
			Text 330, 55 + (10 * z), 40, 10, FormatCurrency(quarterly_output_array(z, 1))
			Text 370, 55 + (10 * z), 40, 10, FormatCurrency(quarterly_output_array(z, 2))
			Text 410, 55 + (10 * z), 40, 10, FormatCurrency(quarterly_output_array(z, 3))
			Text 450, 55 + (10 * z), 40, 10, FormatCurrency(quarterly_output_array(z, 4))	
		NEXT
		Text 15, 105, 60, 10, "Member #, Name"
		Text 110, 105, 40, 10, "Panel, #"
		Text 150, 105, 85, 10, "Employer (Month/Year)"
		Text 330, 105, 30, 10, "Retro"
		Text 370, 105, 30, 10, "Prosp"
		Text 410, 105, 30, 10, "PIC"
		Text 450, 105, 35, 10, "HC Inc Est"
		'Displaying information for each job each month
		FOR i = 1 to num_of_income_panels
			Text 15, (105 + (10 * i)), 120, 10, income_matrix_array(i, 0)
			Text 110, (105 + (10 * i)), 40, 10, income_matrix_array(i, 1)
			Text 150, (105 + (10 * i)), 190, 10, income_matrix_array(i, 2)
			Text 330, (105 + (10 * i)), 30, 10, income_matrix_array(i, 3)
			Text 370, (105 + (10 * i)), 30, 10, income_matrix_array(i, 4)
			Text 410, (105 + (10 * i)), 30, 10, income_matrix_array(i, 5)
			Text 450, (105 + (10 * i)), 30, 10, income_matrix_array(i, 6)
		NEXT
	EndDialog
	CALL navigate_to_MAXIS_screen("CASE", "PERS") 'naving to case pers to see what is currently on what programs. 
	Dialog income_matrix_dlg
		IF ButtonPressed = 0 THEN stopscript
END FUNCTION

'The script......
EMConnect ""

'Navigating to WAGE match
EMSendKey "T"
transmit
EMReadScreen wage, 4, 6, 6
IF wage <> "WAGE" THEN script_end_procedure("These aren't the droids you're looking for.")

CALL write_value_and_transmit("I", 6, 3)
EMReadScreen case_number, 8, 20, 38
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
CALL write_value_and_transmit("IEVP", 20, 71)
EMSendKey "D"
transmit
EMReadScreen all_programs, 10, 7, 13
all_programs = trim(all_programs)
PF3

'Reading IEVS info
CALL write_value_and_transmit("WAGE", 19, 69)
EMReadScreen ievs_name, 30, 4, 25
ievs_name = trim(ievs_name)
EMReadScreen quarterly_wage, 7, 8, 8
EMReadScreen quarter, 1, 8, 16
EMReadScreen ievs_employer, 20, 8, 25
ievs_employer = trim(ievs_employer)
EMReadScreen ievs_year, 4, 8, 19
PF3


ReDim income_matrix_array(20, 6)
CALL income_matrix(income_matrix_array, ievs_name, ievs_employer, quarter, quarterly_wage, ievs_year, all_programs)
CALL check_for_MAXIS(false)
CALL navigate_to_MAXIS_screen("DAIL", "DAIL")

script_end_procedure("")
