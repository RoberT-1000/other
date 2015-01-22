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

EMConnect ""
MAXIS_check_function

'-----Creating the worker array-----
CALL navigate_to_screen("REPT", "USER")
PF5
rept_user_row = 7
DO
	EMReadScreen worker, 7, rept_user_row, 5
	worker = replace(worker, " ", "")
	IF worker <> "" THEN 
		county_array = county_array & worker & " "
		rept_user_row = rept_user_row + 1
	END If
	IF rept_user_row = 19 THEN
		rept_user_row = 7
		PF8
	END IF	
LOOP UNTIL worker = ""

county_array = trim(county_array)
county_array = split(county_array, " ")

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True

'Setting the first 3 col as worker, case number, and name
ObjExcel.Cells(1, 1).Value = "X Number"
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
ObjExcel.Cells(1, 3).Value = "NAME"
ObjExcel.Cells(1, 4).Value = "Under 18 ABAWD Codes"
ObjExcel.Cells(1, 5).Value = "18-24 ABAWD Codes"
ObjExcel.Cells(1, 6).Value = "Over 24 ABAWD Codes"

excel_row = 2
FOR EACH worker IN county_array
	CALL navigate_to_screen("REPT", "ACTV")
	EMWriteScreen worker, 21, 13
	transmit
	EMReadScreen current_worker, 7, 21, 13
	EMReadScreen current_user, 7, 21, 71
	IF current_worker = current_user THEN PF7
	
	rept_actv_row = 7
	DO
		DO
			EMReadScreen last_page, 21, 24, 2
			EMReadScreen case_number, 8, rept_actv_row, 12
				case_number = trim(case_number)
			EMReadScreen case_name, 20, rept_actv_row, 21
			EMReadScreen snap_status, 1, rept_actv_row, 61
			IF snap_status = "A" THEN 
				objExcel.Cells(excel_row, 1).Value = worker
				objExcel.Cells(excel_row, 2).Value = case_number
				objExcel.Cells(excel_row, 3).Value = case_name
				case_array = case_array & case_number & " "
				excel_row = excel_row + 1
			END IF
			rept_actv_row = rept_actv_row + 1
		LOOP UNTIL rept_actv_row = 19 OR case_number = ""
		PF8
	LOOP UNTIL last_page = "THIS IS THE LAST PAGE"
NEXT

case_array = trim(case_array)
case_array = split(case_array, " ")

excel_row = 2
FOR EACH case_number IN case_array
	ABAWD_status = "" 		'clearing variables
	eats_person = ""
	eats_group_members = ""
	eats_group_under18 = ""		
	eats_group_1824 = ""
	eats_group_over24 = ""
	under18 = ""
	from18to24 = ""
	over24 = ""
	
	CALL navigate_to_screen("STAT", "EATS")
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
	
	FOR EACH hh_member IN eats_group_members
		CALL navigate_to_screen("STAT", "MEMB")
		ERRR_screen_check
		
		EMWriteScreen hh_member, 20, 76
		transmit
		
		EMReadScreen age, 3, 8, 76
		IF age = "   " THEN age = 1
		age = age * 1
		
		IF age < 18 THEN
			eats_group_under18 = eats_group_under18 & hh_member & " "
		ELSEIF age >= 18 AND age <= 24 THEN
			eats_group_1824 = eats_group_1824 & hh_member & " "
		ELSEIF age > 24 THEN
			eats_group_over24 = eats_group_over24 & hh_member & " "
		END IF
	NEXT
	
	eats_group_under18 = trim(eats_group_under18)
	eats_group_under18 = split(eats_group_under18, " ")
	eats_group_1824 = trim(eats_group_1824)
	eats_group_1824 = split(eats_group_1824, " ")
	eats_group_over24 = trim(eats_group_over24)
	eats_group_over24 = split(eats_group_over24, " ")
	
	FOR EACH person IN eats_group_under18
		CALL navigate_to_screen("STAT", "WREG")
		ERRR_screen_check
		EMWriteScreen person, 20, 76
		transmit

		EMReadScreen ABAWD_status_code, 2, 13, 50
		under18 = under18 & person & ": " & ABAWD_status_code & ", "
	NEXT
	ObjExcel.Cells(excel_row, 4).Value = under18
	
	FOR EACH person IN eats_group_1824
		CALL navigate_to_screen("STAT", "WREG")
		ERRR_screen_check
		EMWriteScreen person, 20, 76
		transmit
		
		EMReadScreen ABAWD_status_code, 2, 13, 50
		from18to24 = from18to24 & person & ": " & ABAWD_status_code & ", "
	NEXT
	ObjExcel.Cells(excel_row, 5).Value = from18to24
	
	FOR EACH person IN eats_group_over24
		CALL navigate_to_screen("STAT", "WREG")
		ERRR_screen_check
		EMWriteScreen person, 20, 76
		transmit
		
		EMReadScreen ABAWD_status_code, 2, 13, 50
		over24 = over24 & person & ": " & ABAWD_status_code & ", "
	NEXT
	ObjExcel.Cells(excel_row, 6).Value = over24
	
	excel_row = excel_row + 1
NEXT

back_to_SELF
script_end_procedure("Fin")