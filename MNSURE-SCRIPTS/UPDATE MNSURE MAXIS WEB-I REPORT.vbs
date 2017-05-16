'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"

SET req = CreateObject("Msxml2.XMLHttp.6.0") 'Creates an object to get a URL
req.open "GET", url, FALSE	'Attempts to open the URL
req.send 'Sends request

IF req.Status = 200 THEN	'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject") 'Creates an FSO
	Execute req.responseText 'Executes the script code
ELSE	'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox "It's all gone horribly wrong!!!" & _
	vbCr &_
	"URL: " & url
	stopscript
END IF

EMConnect ""

start_time = timer

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("H:\MAXIS MNsure Clients\MNsure MAXIS Clients 2017-05-01.xls")
objExcel.DisplayAlerts = True
objExcel.Worksheets("MAXIS Only").Activate
number_of_cases = 11935			'<< number of rows as of 05/01/2017
'number_of_cases = 10221		'<< number of rows as of 03/01/2017
'number_of_cases = 9928			'<< number of rows as of 02/03/2017
'number_of_cases = 9058			'<< number of rows as of 01/04/2017

On Error Resume Next

' This is the IE shell that displays the status of the report...
Set objExplorer = CreateObject("InternetExplorer.Application")
objExplorer.Navigate "about:blank"   
objExplorer.ToolBar = 0
objExplorer.StatusBar = 0
objExplorer.Width = 800
objExplorer.Height = 100 
objExplorer.Visible = 1             
objExplorer.Document.Title = "Cleaning up report."
objExplorer.Document.Body.InnerHTML = "The script is finding the MAXIS workers."
Wscript.Sleep 1

FOR i = 2 to (number_of_cases + 1)
	' reseting variable values that will be used in the next iteration
	maxis_case_number = objExcel.Cells(i, 7).Value
	x_number = ""
	worker_name = ""
	supervisor = ""
	case_noted_in_MAXIS = FALSE
	PMI_found = false
	cl_PMI = objExcel.Cells(i, 2).Value

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
						EMWriteScreen ("PMI: " & objExcel.Cells(i, 2).Value & ", MNSure ID: " & objExcel.Cells(i, 3).Value & ", Integrated Case: " & objExcel.Cells(i, 1).Value), text_row, 3
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
							EMWriteScreen ("PMI: " & objExcel.Cells(i, 2).Value & ", MNSure ID: " & objExcel.Cells(i, 3).Value & ", Integrated Case: " & objExcel.Cells(i, 1).Value), 5, 3
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
						EMWriteScreen ("PMI: " & objExcel.Cells(i, 2).Value & ", MNSure ID: " & objExcel.Cells(i, 3).Value & ", Integrated Case: " & objExcel.Cells(i, 1).Value), 5, 3
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
	objExcel.Cells(i, 16).Value = x_number
	objExcel.Cells(i, 17).Value = worker_name
	objExcel.Cells(i, 18).Value = supervisor
	IF case_noted_in_MAXIS = TRUE THEN 
		objExcel.Cells(i, 19).Value = "CASE NOTED IN MAXIS"
	ELSE
		objExcel.Cells(i, 19).Value = "COULD NOT CASE NOTE"
	END IF
	
	' generating information about current script run time to display on the report status window
	current_time = timer
	run_time = current_time - start_time
	run_time = FormatNumber(run_time, 2)
	' Updating the status of the report
	objExplorer.Document.Body.InnerHTML = "The script is finding the MAXIS workers. It is " & FormatPercent((i - 1)/number_of_cases) & " complete. Current run time = " & run_time & " seconds. The current row is: " & i
NEXT

' the script has finished running...
msgbox "Success!!"
