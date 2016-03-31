'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - BANKED MONTHS TIKLER.vbs"
start_time = timer

msgbox "File configuration needs to be established for Washington County network mapping. If said configuration has been made, delete content on lines 5 and 6. Otherwise, update line 39 with the correct network location for the Excel file used to store coop information." 
stopscript

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

'Connecting to and checking for MAXIS
EMConnect ""
'CALL check_for_MAXIS(false)

'VARIABLES THAT NEED DECLARING----------------------------------------------------------------------------------------------------
file_path = "Q:\Blue Zone Scripts\Spreadsheets for script use\FSET non-compliance list\non-coop list.xlsx"

'FILESYSTEMOBJECTS FOR SCRIPT----------------------------------------------------------------------------------------------------
Set fso = CreateObject("Scripting.FileSystemObject")

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Grabbing user ID to validate user of script. Only some users are allowed to use this script.
Set objNet = CreateObject("WScript.NetWork") 
user_ID_for_validation = ucase(objNet.UserName)

'Here, you can add a list of individuals (based on their network log-in ID) that are eligible to run this script. These 5 individuals are the 5 Anoka County staff that are authorized to run this script.
'Validating user ID
'If user_ID_for_validation <> "SLCARDA" and _
'	user_ID_for_validation <> "EABUELOW" and _
'	user_ID_for_validation <> "CDPOTTER" and _
'	user_ID_for_validation <> "RAKALB" AND _
'	user_ID_for_validation <> "PHBROCKM" _
'	then script_end_procedure("User " & user_ID_for_validation & " is not authorized to use this script. To be added to the allowed users' group, email the script administrator, and include the user ID indicated. Thank you!")

'Checks to make sure the file exists. If it doesn't the script will exit.
If fso.FileExists(file_path) = False then script_end_procedure("''No show list'' not found. The list should be saved at " & file_path & ". Check to make sure the file is there and try again.")

'LOADING EXCEL
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Open(file_path) 
objExcel.DisplayAlerts = False 'Set this to false to make alerts go away. This is necessary in production.

FOR excel_row = 2 to 300
	'Grabs the MAXIS case number
	'Note from 03/30/2016
	'Changing the locations. When we ran the non-coop check, banked_coop was found on column 1, case number on column 3, and cl_pmi on column 4
	case_number = objExcel.Cells(excel_row, 2).Value
	cl_pmi = trim(objExcel.Cells(excel_row, 3).Value)
	IF len(cl_pmi) <> 8 THEN 
		DO
			cl_pmi = "0" & cl_pmi
		LOOP UNTIL len(cl_pmi) = 8
	END IF
	banked_coop = UCASE(objExcel.Cells(excel_row, 4).Value)
	privileged_check = ""
	memb_num = ""
	county_code = ""
	pmi_found = FALSE
	IF banked_coop <> "" THEN 
		IF case_number <> "" THEN 
			IF coop_mode = TRUE THEN 
				keep_snap_open_month = DateAdd("M", 1, date)
				IF len(DatePart("M", keep_snap_open_month)) = 1 THEN 
					keep_snap_open_month = "0" & DatePart("M", keep_snap_open_month) & "/" & Right(DatePart("YYYY", keep_snap_open_month), 2)
				ELSE
					keep_snap_open_month = DatePart("M", keep_snap_open_month) & "/" & Right(DatePart("YYYY", keep_snap_open_month), 2)
				END IF
				IF banked_coop = "YES" THEN 
					back_to_SELF
					EMWriteScreen "STAT", 16, 43
					EMWriteScreen "________", 18, 43
					EMWriteScreen case_number, 18, 43
					EMWriteScreen "MEMB", 21, 70
					transmit
					
					EMReadScreen privileged_check, 70, 24, 2
					IF InStr(privileged_check, "PRIVILEGE") <> 0 THEN 
						objExcel.Cells(excel_row, 5).Value = "PRIVILEGED"
					ELSE
						EMReadScreen county_code, 2, 21, 8
						IF county_code <> "82" THEN 
							objExcel.Cells(excel_row, 5).Value = "OUT OF WASHINGTON COUNTY"
						ELSE						
							DO
								EMReadScreen maxis_pmi, 8, 4, 46
								maxis_pmi = trim(maxis_pmi)
								IF len(maxis_pmi) <> 8 THEN 
									DO
										maxis_pmi = "0" & maxis_pmi
									LOOP UNTIL len(maxis_pmi) = 8
								END IF
								'msgbox cl_pmi & vbCr & maxis_pmi
								IF maxis_pmi = cl_pmi THEN 
									EMReadScreen memb_num, 2, 4, 33
									pmi_found = TRUE
									EXIT DO
								ELSE 
									transmit
									EMReadScreen enter_a_valid_command, 21, 24, 2
									IF enter_a_valid_command = "ENTER A VALID COMMAND" THEN EXIT DO
								END IF
							LOOP
							'msgbox pmi_found
							IF pmi_found = TRUE THEN 
								CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
								CALL write_variable_in_TIKL("MEMB: " & memb_num & " cooperating with Employment Services. Review ABAWD/Banked Months.") 
								transmit
								PF3
							ELSE
								objExcel.Cells(excel_row, 5).Value = "PMI NOT FOUND ON THIS CASE"
							END IF
						END IF
					END IF
				END IF
			END IF
			IF noncoop_mode = TRUE THEN 
				close_snap_date = DateAdd("M", 1, date)
				IF len(DatePart("M", close_snap_date)) = 1 THEN 
					close_snap_date = "0" & DatePart("M", close_snap_date) & "/01/" & DatePart("YYYY", close_snap_date)
				ELSE
					close_snap_date = DatePart("M", close_snap_date) & "/01/" & DatePart("YYYY", close_snap_date)
				END IF
				IF banked_coop = "NO" THEN 
					back_to_SELF
					EMWriteScreen "STAT", 16, 43
					EMWriteScreen "________", 18, 43
					EMWriteScreen case_number, 18, 43
					EMWriteScreen "MEMB", 21, 70
					transmit
					
					EMReadScreen privileged_check, 70, 24, 2
					IF InStr(privileged_check, "PRIVILEGE") <> 0 THEN 
						objExcel.Cells(excel_row, 5).Value = "PRIVILEGED"
					ELSE
					
						EMReadScreen county_code, 2, 21, 8
						IF county_code <> "82" THEN 
							objExcel.Cells(excel_row, 5).Value = "OUT OF WASHINGTON COUNTY"
						ELSE
							DO
								EMReadScreen maxis_pmi, 8, 4, 46
								maxis_pmi = trim(maxis_pmi)
								IF maxis_pmi = cl_pmi THEN 
									EMReadScreen memb_num, 2, 4, 33
									pmi_found = TRUE
									EXIT DO
								ELSE 
									transmit
									EMReadScreen enter_a_valid_command, 21, 24, 2
									IF enter_a_valid_command = "ENTER A VALID COMMAND" THEN EXIT DO
								END IF
							LOOP
						
							IF pmi_found = TRUE THEN 
								CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
								CALL write_variable_in_TIKL("MEMB: " & memb_num & " not cooperating with Employment Services this month.") 
								transmit
								PF3
							ELSE
								objExcel.Cells(excel_row, 5).Value = "PMI NOT FOUND ON THIS CASE"
							END IF
						END IF
					END IF
				END IF
			END IF
		ELSE
			EXIT FOR
		END IF
	END IF
NEXT

script_end_procedure("TIKLs created.")
