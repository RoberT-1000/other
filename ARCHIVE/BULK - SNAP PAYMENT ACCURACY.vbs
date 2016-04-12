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

BeginDialog snap_payment_accuracy_dialog, 0, 0, 321, 110, "SNAP Payment Accuracy"
  CheckBox 15, 20, 80, 10, "ABAWD Status", abawd_check
  CheckBox 15, 35, 85, 10, "SNAP w/ GA/RCA", ga_rca_check
  DropListBox 205, 15, 85, 15, "REPT/ACTV"+chr(9)+"REPT/PND2", screen_to_use
  EditBox 210, 35, 50, 15, worker_number
  CheckBox 120, 60, 195, 10, "Check here to run the script for all users in your agency.", all_worker_check
  ButtonGroup ButtonPressed
    OkButton 210, 85, 50, 15
    CancelButton 260, 85, 50, 15
  GroupBox 5, 10, 105, 90, "SNAP Reports"
  Text 120, 15, 70, 10, "Select Source:"
  Text 120, 40, 85, 10, "Enter worker X number(s)"
EndDialog

EMConnect ""

benefit_month = datepart("M", dateadd("M", 1, date))
IF len(benefit_month) <> 2 THEN benefit_month = "0" & benefit_month
benefit_year = datepart("YYYY", dateadd("M", 1, date))
benefit_year = right(benefit_year, 2)

back_to_SELF
EMWriteScreen benefit_month, 20, 43
EMWriteScreen benefit_year, 20, 46

DO
	DIALOG snap_payment_accuracy_dialog
		IF ButtonPressed = 0 THEN stopscript
LOOP UNTIL (worker_number = "" AND all_worker_check = 1) OR (all_worker_check = 0 AND worker_number <> "")

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True

'Setting the first 3 col as worker, case number, and name
ObjExcel.Cells(1, 1).Value = "X Number"
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
ObjExcel.Cells(1, 3).Value = "NAME"

col_to_use = 4
IF abawd_check = 1 THEN
	ObjExcel.Cells(1, col_to_use).Value = "Under 18 ABAWD Status"
	under18_col = col_to_use
	col_to_use = col_to_use + 1
	ObjExcel.Cells(1, col_to_use).Value = "18 - 50 ABAWD Status"
	from18to24_col = col_to_use
	col_to_use = col_to_use + 1
	ObjExcel.Cells(1, col_to_use).Value = "Over 50 ABAWD Status"
	over50_col = col_to_use
	col_to_use = col_to_use + 1
END IF
IF ga_rca_check = 1 THEN
	ObjExcel.Cells(1, col_to_use).Value = "GA/RCA FIAT'ing Error?"	
	ga_rca_col = col_to_use
	col_to_use = col_to_use + 1
END IF
	

IF all_worker_check = 1 THEN
	CALL navigate_to_screen("REPT", "USER")
	PF5

	rept_user_row = 7
	DO
		EMReadScreen worker_number, 7, rept_user_row, 5
		worker_number = trim(worker_number)
		IF worker_number <> "" THEN worker_array = worker_array & worker_number & " "
		rept_user_row = rept_user_row + 1
		IF rept_user_row = 19 THEN
			rept_user_row = 7
			PF8
		END IF
		EMReadScreen last_page, 21, 24, 2
	LOOP UNTIL worker_number = "" OR last_page = "THIS IS THE LAST PAGE"

	worker_array = trim(worker_array)
	worker_array = split(worker_array)
ELSE
	worker_array = split(worker_number, ", ")
END IF


excel_row = 2
case_array = ""
FOR EACH worker IN worker_array
	case_number = ""
	IF screen_to_use = "REPT/ACTV" THEN
		CALL navigate_to_screen("REPT", "ACTV")
		EMWriteScreen worker, 21, 13
		transmit
	
		CALL find_variable("User: ", current_user, 7)
		IF ucase(worker) = ucase(current_user) THEN PF7
	
		rept_actv_row = 7
		DO
			DO
				EMReadScreen last_page, 21, 24, 2
				EMReadScreen case_number, 8, rept_actv_row, 12
				case_number = trim(case_number)
				EMReadScreen snap_status, 1, rept_actv_row, 61
				EMReadScreen cash_status, 1, rept_actv_row, 54
				EMReadScreen client_name, 20, rept_actv_row, 21
				IF (abawd_check = 1 AND snap_status = "A") OR (ga_rca_check = 1 AND snap_status = "A" AND cash_status <> " ") THEN
					case_array = case_array & case_number & " "
					objExcel.Cells(excel_row, 1).Value = worker
					objExcel.Cells(excel_row, 2).Value = case_number
					objExcel.Cells(excel_row, 3).Value = client_name
					excel_row = excel_row + 1
				END IF
				rept_actv_row = rept_actv_row + 1
			LOOP UNTIL rept_actv_row = 19
				PF8
				rept_actv_row = 7
		LOOP UNTIL case_number = "" OR last_page = "THIS IS THE LAST PAGE"
	ELSEIF screen_to_use = "REPT/PND2" THEN
		back_to_SELF
		CALL navigate_to_screen("REPT", "PND2")
		EMWriteScreen worker, 21, 13
		transmit

		EMReadScreen no_content, 6, 3, 74
		IF no_content <> "0 Of 0" THEN 
			rept_pnd2_row = 7
			DO
				DO
					EMReadScreen last_page, 21, 24, 2
					EMReadScreen case_number, 8, rept_pnd2_row, 5
					case_number = trim(case_number)
					EMReadScreen client_name, 20, rept_pnd2_row, 16
					IF case_number <> "" THEN 
						case_array = case_array & case_number & " "
						objExcel.Cells(excel_row, 1).Value = worker
						objExcel.Cells(excel_row, 2).Value = case_number
						objExcel.Cells(excel_row, 3).Value = client_name
						excel_row = excel_row + 1
					END IF
					rept_pnd2_row = rept_pnd2_row + 1
				LOOP UNTIL case_number = "" OR rept_pnd2_row = 19
				PF8
				rept_pnd2_row = 7
			LOOP UNTIL last_page = "THIS IS THE LAST PAGE"
		END IF
	END IF
NEXT

case_array = trim(case_array)
case_array = split(case_array)

excel_row = 2
'The script will go to the following places, depending on which reports have been requested...
'	IF the script is checking for GA & RCA being FIAT'd into the SNAP budget, the script will navigate to CASE/CURR to determine which cash program to check
'	THE script will navigate to ELIG/FS and navigate to the most recently approved version.
'		IF the script is checking ABAWD codes, it will create a list of eligible household members on FSPR
'		IF the script is looking for the GA/RCA amount FIAT'd, it will go to FSB1 to find the PA Amount FIAT'd. It will then go into either ELIG/GA or ELIG/RCA to determine the amount approved for the next month.
'	IF the script is checking ABAWD codes, it will generate 3 lists to check the ABAWD statuses of eligible household members based on age. 16-17 y/o is list 1, 18-49 y/o is list 2, and 50+ y/o is list 3.
'		IF the script encounters an invalid ABAWD code (such as "__" or "09" or "02" (for individuals aged 18 or older) or "03" (for individuals aged 49 or younger) then it will highlight that cell in yellow.
'	IF the script is checking GA/RCA fiating, it will navigate to STAT/REVW to determine if the discrepancy is allowable considering the case is in a review month.
'	The script will then update Excel with the information.

FOR EACH case_number IN case_array
	'Clearing the variables
	ABAWD_status = ""
	hh_array = ""
	under18_array = ""
	from18to50_array = ""
	over50_array = ""
	ga_status = ""
	ga_amount = ""
	rca_status = ""
	rca_amount = ""
	cash_prog = ""
	pa_amount = ""
	under18 = ""
	from18to50 = ""
	over50 = ""

	IF ga_rca_check = 1 THEN 
		CALL navigate_to_screen("CASE", "CURR")
		CALL find_variable("GA: ", ga_status, 6)
		IF ga_status = "ACTIVE" OR ga_status = "APP CL" THEN
			cash_prog = "GA"
		ELSE
			CALL find_variable("RCA: ", rca_status, 6)
			IF rca_status = "ACTIVE" OR rca_status = "APP CL" THEN cash_prog = "RCA"
		END IF
	END IF

	CALL navigate_to_screen("ELIG", "FS")

	EMReadScreen approved, 8, 3, 3
	IF approved <> "APPROVED" THEN 
		EMReadScreen version, 2, 2, 12
		version = version * 1
		version = version - 1
		IF len(version) <> 2 THEN version = "0" & version
		EMWriteScreen version, 19, 78
		transmit
	END IF

	IF abawd_check = 1 THEN
		fspr_row = 7
		DO
			DO
				EMReadScreen hh_memb, 2, fspr_row, 10
				EMReadScreen elig_status, 4, fspr_row, 57
				IF elig_status = "ELIG" THEN 
					hh_array = hh_array & hh_memb & " "
				END IF
				fspr_row = fspr_row + 1
			LOOP UNTIL fspr_row = 18
			PF8
			EMReadScreen no_more_members, 15, 24, 5
			fspr_row = 7
		LOOP UNTIL no_more_members = "NO MORE MEMBERS"

		hh_array = trim(hh_array)
		hh_array = split(hh_array)
	END IF

	IF ga_rca_check = 1 THEN
		EMWriteScreen "FSB1", 19, 70
		transmit
	
		CALL find_variable("PA Grants..............$", pa_amount, 10)
		pa_amount = replace(pa_amount, "_", "")
		pa_amount = trim(pa_amount)
		IF pa_amount = "" THEN pa_amount = "0.00"

		IF cash_prog = "GA" THEN
			CALL navigate_to_screen("ELIG", "GA")
			EMReadScreen approved, 8, 3, 3
			EMReadScreen version, 2, 2, 12
			version = trim(version)
			version = version - 1
			IF len(version) <> 2 THEN version = "0" & version
			IF approved <> "APPROVED" THEN 
				EMWriteScreen version, 20, 78
				transmit
			END IF
			EMWriteScreen "GASM", 20, 70
			transmit
				CALL find_variable("Monthly Grant............$", ga_amount, 9)
			ga_amount = trim(ga_amount)
			IF pa_amount <> ga_amount THEN
				objExcel.Cells(excel_row, ga_rca_col).Value = ("Yes, GA. SNAP Budg = " & pa_amount & "; GA Amount = " & ga_amount)
				objExcel.Cells(excel_row, ga_rca_col).Interior.ColorIndex = 6
			ELSEIF pa_amount = ga_amount THEN
				objExcel.Cells(excel_row, ga_rca_col).Value = ("Budgeted for SNAP: " & pa_amount & "; GA Amount: " & ga_amount)
			END IF
		ELSEIF cash_prog = "RCA" THEN
			CALL navigate_to_screen("ELIG", "RCA")
			EMReadScreen approved, 8, 3, 3
			EMReadScreen version, 2, 2, 12
			version = trim(version)
			version = version - 1
			IF len(version) <> 2 THEN version = "0" & version
			IF approved <> "APPROVED" THEN 
				EMWriteScreen version, 19, 78
				transmit
			END IF
				EMWriteScreen "RCSM", 19, 70
			transmit
		
			CALL find_variable("Grant Amount..............$", rca_amount, 10)
			rca_amount = trim(rca_amount)
			IF pa_amount <> rca_amount THEN 
				objExcel.Cells(excel_row, ga_rca_col).Value = ("Yes, RCA. SNAP Budg = " & pa_amount & "; RCA Amount = " & rca_amount)
				objExcel.Cells(excel_row, ga_rca_col).Interior.ColorIndex = 6
			ELSEIF pa_amount = rca_amount THEN
				objExcel.Cells(excel_row, ga_rca_col).Value = ("Budgeted for SNAP: " & pa_amount & "; RCA Amount: " & rca_amount)
			END IF
		END IF
	END IF

	IF abawd_check = 1 THEN 
		FOR EACH hh_memb IN hh_array
			IF hh_memb <> "  " THEN 
				CALL navigate_to_screen("STAT", "MEMB")
				ERRR_screen_check
	
				IF hh_memb <> "01" THEN
					EMWriteScreen hh_memb, 20, 79
					transmit
				END IF
				
				CALL find_variable("Age: ", cl_age, 3)
				cl_age = trim(cl_age)
				IF cl_age = "" THEN cl_age = "1"
	
				cl_age = cl_age * 1
					
				IF cl_age >= 50 THEN
					over50_array = over50_array & hh_memb & " "
				ELSEIF cl_age > 17 AND cl_age < 50 THEN
					from18to50_array = from18to50_array & hh_memb & " "
				ELSEIF cl_age > 15 AND cl_age < 18 THEN
					under18_array = under18_array & hh_memb & " "
				END IF
			END IF
		NEXT
			
		under18_array = trim(under18_array)
		under18_array = split(under18_array, " ")

		from18to50_array = trim(from18to50_array)
		from18to50_array = split(from18to50_array, " ")

		over50_array = trim(over50_array)
		over50_array = split(over50_array, " ")

		FOR EACH hh_memb IN under18_array
			EMWriteScreen "WREG", 20, 71
			EMWriteScreen hh_memb, 20, 79
			transmit

			EMReadScreen abawd_code, 2, 13, 50
			IF abawd_code = "__" OR abawd_code = "03" OR abawd_code = "09" THEN objExcel.Cells(excel_row, over50_col).Interior.ColorIndex = 6
			under18 = under18 & hh_memb & "; " & abawd_code & ", "
			ObjExcel.Cells(excel_row, under18_col).Value = under18
		NEXT

		FOR EACH hh_memb IN from18to50_array
			EMWriteScreen "WREG", 20, 71
			EMWriteScreen hh_memb, 20, 79
			transmit

			EMReadScreen abawd_code, 2, 13, 50
			IF abawd_code = "__" OR abawd_code = "02" OR abawd_code = "03" OR abawd_code = "09" THEN objExcel.Cells(excel_row, over50_col).Interior.ColorIndex = 6
			from18to50 = from18to50 & hh_memb & "; " & abawd_code & ", "
			ObjExcel.Cells(excel_row, from18to24_col).Value = from18to50
		NEXT
		
		FOR EACH hh_memb IN over50_array
			EMWriteScreen "WREG", 20, 71
			EMWriteScreen hh_memb, 20, 79
			transmit

			EMReadScreen abawd_code, 2, 13, 50
			IF abawd_code = "__" OR abawd_code = "02" OR abawd_code = "09" THEN objExcel.Cells(excel_row, over50_col).Interior.ColorIndex = 6
			over50 = over50 & hh_memb & "; " & abawd_code & ", "
			objExcel.Cells(excel_row, over50_col).Value = over50
		NEXT
	END IF

	IF ga_rca_check = 1 THEN 
		CALL navigate_to_screen("STAT", "REVW")
		ERRR_screen_check
		EMReadScreen cash_revw_date, 8, 9, 37
		EMReadScreen snap_revw_date, 8, 9, 57
		bene_date = benefit_month & "/" & benefit_year
		cash_revw_date = replace(cash_revw_date, " 01 ", "/")
		snap_revw_date = replace(snap_revw_date, " 01 ", "/")
		IF bene_date = cash_revw_date OR bene_date = snap_revw_date THEN objExcel.Cells(excel_row, ga_rca_col).Value = "REVW MONTH"
	END IF

	excel_row = excel_row + 1
NEXT

script_end_procedure("Done")
