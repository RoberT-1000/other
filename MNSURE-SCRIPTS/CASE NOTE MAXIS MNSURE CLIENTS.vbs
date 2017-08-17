'LOADING GLOBAL VARIABLES--------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Public assistance script files\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			'If the connection to GitHub is severed, we want users to run the backup
			github_live = FALSE
			SET run_backup_script = CreateObject("Scripting.FileSystemObject")
			SET run_backup_script_command = run_backup_script.OpenTextFile("Q:\Blue Zone Scripts\FUNCTIONS LIBRARY.vbs")
			backup_script = run_backup_script_command.ReadAll
			run_backup_script_command.Close
			Execute backup_script
		END IF
	ELSE
		github_live = FALSE
		FuncLib_URL = "Q:\Blue Zone Scripts\FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

EMConnect ""

start_time = timer

'red rows are privileged
'green rows are cases where the PMI could not be found
'yellow rows are cases outside of X102

'calling function for opening Excel file...
CALL anoka_file_selection_system_dialog(excel_file_path, ".xlsx", ".xls")

SET objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
SET objWorkbook = objExcel.Workbooks.Add(excel_file_path)
objExcel.DisplayAlerts = FALSE
'objExcel.Worksheets("MAXIS Only").Activate

' =======================================================================================
' =======================================================================================
' ============== THIS SECTION OF CODE HAS BEEN COMMENTED OUT -- IT REQURIES =============
' ============== TOO MUCH MANUAL SUPPORT TO RUN WITH THE PROGRESS BAR. ==================
' ============== IF YOU ARE INTERESTED IN HAVING THE PROGRESS BAR, YOU NEED =============
' ============== TO SPECIFY THE NUMBER OF CASES. ========================================
' =======================================================================================
' =======================================================================================
	'On Error Resume Next
	' This is the IE shell that displays the status of the report...
	'Set objExplorer = CreateObject("InternetExplorer.Application")
	'objExplorer.Navigate "about:blank"   
	'objExplorer.ToolBar = 0
	'objExplorer.StatusBar = 0
	'objExplorer.Width = 900
	'objExplorer.Height = 100 
	'objExplorer.Visible = 1             
	'objExplorer.Document.Title = "Determining Number of Cases"
	'objExplorer.Document.Body.InnerHTML = "The script is determining the number of cases. Please standby..."
	'Wscript.Sleep 1
	
	
	'number_of_cases = 12862		'<< number of rows as of 07/05/2017
	'number_of_cases = 7792			'<< number of rows as of 06/01/2017	Why the drop off?
	'number_of_cases = 11935		'<< number of rows as of 05/01/2017
	'number_of_cases = 10221		'<< number of rows as of 03/01/2017
	'number_of_cases = 9928			'<< number of rows as of 02/03/2017
	'number_of_cases = 9058			'<< number of rows as of 01/04/2017
' =======================================================================================
' =======================================================================================

'FOR i = 2 to (number_of_cases + 1)
i = 2
do
	' columns on the un-edited spreadsheet...
	' 1 = Service Location
	' 2 = County of Residence
	' 3 = Integrated Case
	' 4 = Recipient ID
	' 5 = MNsure ID
	' 6 = FirstName 
	' 7 = MiddleInitial 
	' 8 = LastName 
	' 9 = MajorProgram 
	' 10 = BeginDate 
	' 11 = End Date
	' 12 = MAXISCase Number
	' 13 = Program 
	' 14 = Status 
	' 15 = Birth Date 
	' 16 = Gender 
	' 17 = SSN 
	' 18 = SMI Number 
	' 19 = Address 1 
	' 20 = Address 2 
	' 21 = City 
	' 22 = State 
	' 23 = Zip Code
	' 24 = ADDED BY SCRIPT >> X102 Number
	' 25 = ADDED BY SCRIPT >> Worker Name
	' 26 = ADDED BY SCRIPT >> Supervisor
	' 27 = ADDED BY SCRIPT >> CASE NOTED?
	'=============================================
	' reseting variable values that will be used in the next iteration
	maxis_case_number = objExcel.Cells(i, 12).Value
	maxis_case_number = replace(maxis_case_number, " ", "")
	if maxis_case_number <> "" then 
		cl_PMI = objExcel.Cells(i, 4).Value	
		mnsure_id = objExcel.Cells(i, 5).Value
		integrated_case = objExcel.Cells(i, 3).Value
		x_number = ""
		worker_name = ""
		supervisor = ""
		case_noted_in_MAXIS = FALSE
		PMI_found = false
	
		back_to_SELF
		
		'The script starts by trying to get into STAT/MEMB...first to grab the worker's X102 number, their name, and their supervisor's name...
		CALL navigate_to_MAXIS_screen("STAT", "MEMB")
		
		'...checking to see that the script was able to make it off of SELF...
		EMReadScreen at_self, 4, 2, 50
		IF at_self = "SELF" THEN 
			CALL find_variable("PRIVILEGED WORKER: ", x_number, 7)
			supervisor = "PRIVILEGED"
		ELSE
			'...checking to make sure that the script got past ERRR...
			DO
				EMReadScreen at_MEMB, 4, 2, 48
				IF at_MEMB <> "MEMB" THEN transmit
			LOOP UNTIL at_MEMB = "MEMB"
			
			'...reading for the X102 number, the worker's name, and the supervisor's name...
			EMReadScreen x_number, 7, 21, 21
			EMSetCursor 21, 21
			PF1
			EMReadScreen worker_name, 20, 19, 10
			worker_name = trim(worker_name)
			EMReadScreen supervisor, 20, 22, 16
			supervisor = trim(supervisor)
			transmit
			
			'...checking to look for the PMI on that specific case...		
			DO
				EMReadScreen PMI_num, 8, 4, 46
				PMI_num = trim(PMI_num)
				DO
					IF len(PMI_num) <> 8 THEN PMI_num = "0" & PMI_num
				LOOP UNTIL len(PMI_num) = 8
				
				IF cl_PMI = PMI_num THEN
					PMI_found = true
					EXIT DO
				ELSE
					transmit
				END IF
				EMReadScreen enter_a_valid_command, 21, 24, 2
			LOOP UNTIL enter_a_valid_command = "ENTER A VALID COMMAND"
			
			'If the case is not in X102 then the script will not case note. We are also ONLY going to continue on this case when the PMI is found.
			IF UCASE(LEFT(x_number, 4)) = "X102" AND UCASE(x_number) <> "X102CLS" AND PMI_found = TRUE THEN
				'Case noting that the client is active in both systems
				PF4
				'looking for affiliated MNSure Case noting
				case_note_row = 5
				DO
					case_note_header = ""
					case_note_date = ""
					EMReadScreen case_note_header, 30, case_note_row, 25
					IF case_note_header = "~~~ AFFILIATED MNSURE CASE ~~~" THEN 
						EMReadScreen case_note_date, 8, case_note_row, 6
						IF CDate(case_note_date) = DATE THEN 
							CALL write_value_and_transmit("X", case_note_row, 3)
							PF9
							text_row = 5
							DO
								case_note_text = ""
								EMReadScreen case_note_text, 70, text_row, 3
								case_note_text = trim(case_note_text)
								IF case_note_text <> "" THEN 
									text_row = text_row + 1
									IF text_row = 18 THEN 
										PF8
										text_row = 4
									END IF
								END IF
							LOOP UNTIL case_note_text = ""
							EMWriteScreen ("PMI: " & cl_PMI & ", MNSure ID: " & mnsure_id & ", Integrated Case: " & integrated_case), text_row, 3
							transmit
							PF3
							case_noted_in_MAXIS = TRUE
							EXIT DO
						ELSE
							case_note_row = case_note_row + 1
							IF case_note_row = 19 THEN 
								'Creating new case note because one has not been found on this date
								PF9
								EMWriteScreen "~~~ AFFILIATED MNSURE CASE ~~~", 4, 3
								EMWriteScreen ("PMI: " & cl_PMI & ", MNSure ID: " & mnsure_id & ", Integrated Case: " & integrated_case), 5, 3
								transmit
								PF3
								case_noted_in_MAXIS = TRUE
								EXIT DO
							END IF
						END IF
					ELSE
						case_note_row = case_note_row + 1
						IF case_note_row = 19 THEN 
							'Creating new case note because one has not been found on this date
							PF9
							EMWriteScreen "~~~ AFFILIATED MNSURE CASE ~~~", 4, 3
							EMWriteScreen ("PMI: " & cl_PMI & ", MNSure ID: " & mnsure_id & ", Integrated Case: " & integrated_case), 5, 3
							transmit
							PF3
							case_noted_in_MAXIS = TRUE
							EXIT DO
						END IF
					END IF
				LOOP 
			END IF
		END IF
		
		' adding the worker's X102, the worker's name, and the worker's supervisor's name to the Excel file
		objExcel.Cells(i, 24).Value = x_number
		objExcel.Cells(i, 25).Value = worker_name
		objExcel.Cells(i, 26).Value = supervisor
		IF case_noted_in_MAXIS = TRUE THEN 
			objExcel.Cells(i, 27).Value = "CASE NOTED IN MAXIS"
		ELSE
			objExcel.Cells(i, 27).Value = "COULD NOT CASE NOTE"
		END IF
		
		' generating information about current script run time to display on the report status window
		'current_time = timer
		'run_time = current_time - start_time
		'run_time = FormatNumber(run_time, 2)
		' Updating the status of the report
		'objExplorer.Document.Body.InnerHTML = "The script is finding the MAXIS workers. It is " & FormatPercent((i - 1)/number_of_cases) & " complete. Current run time = " & run_time & " seconds. The current row is: " & i
		
		' saving the file every 250 entries	
		if (i Mod 1000) = 0 THEN objWorkbook.Saveas excel_file_path
	end if
	i = i + 1
loop until objExcel.Cells(i, 8).Value = ""

' the script has finished running...
msgbox "Success!!"
