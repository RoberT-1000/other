'Gathering stats==============================================================================
name_of_script = "BULK - MEMO FROM LIST.vbs"
start_time = timer

''LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
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

'Dialogs
'>>>>>Main dlg<<<<<
BeginDialog main_menu, 0, 0, 201, 65, "MEMO from List"
  DropListBox 5, 40, 80, 10, "Manual Entry"+chr(9)+"REPT/ACTV"+chr(9)+"Excel File", run_mode
  ButtonGroup ButtonPressed
    OkButton 90, 40, 50, 15
    CancelButton 140, 40, 50, 15
  Text 10, 10, 185, 25, "Please select a run mode for the script. You can either enter the case numbers manually, from REPT/ACTV, or from an Excel file..."
EndDialog

'>>>>> Function to build dlg for manual entry <<<<<
FUNCTION build_manual_entry_dlg(case_number_array, memo_text)
	'Array for all case numbers
	'This was chosen over building a dlg with 50 variables
	REDim all_cases_array(50, 0)
	
	BeginDialog man_entry_dlg, 0, 0, 331, 310, "Enter MAXIS case numbers"
		Text 10, 15, 140, 10, "Enter MAXIS case numbers below..."
		dlg_row = 30
		dlg_col = 10
		FOR i = 1 TO 50
			EditBox dlg_col, dlg_row, 55, 15, all_cases_array(i, 0)
			dlg_row = dlg_row + 20
			IF dlg_row = 230 THEN 
				dlg_row = 30
				dlg_col = dlg_col + 65
			END IF
		NEXT
		Text 10, 240, 90, 10, "Enter your MEMO text..."
		EditBox 10, 255, 310, 15, memo_text
		ButtonGroup ButtonPressed
			OkButton 220, 290, 50, 15
			CancelButton 270, 290, 50, 15
	EndDialog
	
	'Calling the dlg within the function
	DO 
		'err_msg handling
		err_msg = ""
		DIALOG man_entry_dlg
			cancel_confirmation
			FOR i = 1 TO 50
				all_cases_array(i, 0) = replace(all_cases_array(i, 0), " ", "")
				IF all_cases_array(i, 0) <> "" THEN 
					IF len(all_cases_array(i, 0)) > 8 THEN err_msg = err_msg & vbCr & "* Case number " & all_cases_array(i, 0) & " is too long to be a valid MAXIS case number."
					IF isnumeric(all_cases_array(i, 0)) = FALSE THEN err_msg = err_msg & vbCr & "* Case number " & all_cases_array(i, 0) & " contains alphabetic characters. These are not valid."
				END IF
			NEXT
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
	
	'building the array
	case_number_array = ""
	FOR i = 1 TO 50
		IF all_cases_array(i, 0) <> "" THEN case_number_array = case_number_array & all_cases_array(i, 0) & "~~~"
	NEXT
END FUNCTION

'>>>>>DLG for Excel mode<<<<<
BeginDialog memo_from_excel_dlg, 0, 0, 256, 135, "MEMO Information"
  EditBox 220, 10, 25, 15, excel_col
  EditBox 65, 30, 40, 15, excel_row
  EditBox 190, 30, 40, 15, end_row
  EditBox 10, 70, 235, 15, memo_text
  ButtonGroup ButtonPressed
    OkButton 130, 115, 55, 15
    CancelButton 190, 115, 60, 15
  Text 10, 15, 205, 10, "Please enter the column containing the MAXIS case numbers..."
  Text 10, 35, 50, 10, "Row to start..."
  Text 135, 35, 50, 10, "Row to end..."
  Text 10, 55, 230, 10, "Please enter your MEMO text. Separate new lines with semi-colons..."
EndDialog

'>>>>> THE DLG for REPT/ACTV mode<<<<<
BeginDialog worker_number_dlg, 0, 0, 231, 60, "Enter worker number"
  EditBox 145, 10, 65, 15, worker_number
  ButtonGroup ButtonPressed
    OkButton 65, 35, 50, 15
    CancelButton 115, 35, 50, 15
  Text 10, 15, 130, 10, "Please enter the 7-digit worker number:"
EndDialog

'----------FUNCTIONS----------
'-----This function needs to be added to the FUNCTIONS FILE-----
'>>>>> This function converts the letter for a number so the script can work with it <<<<<
FUNCTION convert_excel_letter_to_excel_number(excel_col)
	IF isnumeric(excel_col) = FALSE THEN 
		alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		excel_col = ucase(excel_col)
		IF len(excel_col) = 1 THEN 
			excel_col = InStr(alphabet, excel_col)
		ELSEIF len(excel_col) = 2 THEN 
			excel_col = (26 * InStr(alphabet, left(excel_col, 1))) + (InStr(alphabet, right(excel_col, 1)))
		END IF
	ELSE
		excel_col = CInt(excel_col)
	END IF
END FUNCTION

'-------THIS FUNCTION ALLOWS THE USER TO PICK AN EXCEL FILE---------
Function BrowseForFile()
    Dim shell : Set shell = CreateObject("Shell.Application")
    Dim file : Set file = shell.BrowseForFolder(0, "Choose a file:", &H4000, "Computer")
	IF ButtonPressed = -1 THEN script_end_procedure("The script will end.")
    BrowseForFile = file.self.Path
End Function

'The script===========================
EMConnect ""

CALL check_for_MAXIS(true)

'>>>>> loading the main dialog <<<<<
DIALOG main_menu
	IF ButtonPressed = 0 THEN stopscript
	'>>>>> the script has different ways of building case_number_array
	IF run_mode = "Manual Entry" THEN 
		CALL build_manual_entry_dlg(case_number_array, memo_text)
		
	ELSEIF run_mode = "REPT/ACTV" THEN 
		'script_end_procedure("This mode is not yet supported.")
		CALL find_variable("User: ", worker_number, 7)
		DO
			err_msg = ""
			DIALOG worker_number_dlg
				cancel_confirmation
				worker_number = trim(worker_number)
				IF worker_number = "" THEN err_msg = err_msg & vbCr & "* You must enter a worker number."
				IF len(worker_number) <> 7 THEN err_msg = err_msg & vbCr & "* Your worker number must be 7 characters long."
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL err_msg = ""
		
		CALL check_for_MAXIS(false)
		
		'Checking that case number is blank so as to get a full REPT/ACTV
		CALL find_variable("Case Nbr: ", case_number, 8)
		case_number = replace(case_number, "_", " ")
		case_number = trim(case_number)
		IF case_number <> "" THEN 
			back_to_SELF
			EMWriteScreen "________", 18, 43
		END IF	
		'Checking that MAXIS is not already in REPT/ACTV so as to get a full REPT/ACTV
		EMReadScreen at_REPT_ACTV, 4, 2, 48
		IF at_REPT_ACTV = "ACTV" THEN back_to_SELF
		
		CALL navigate_to_MAXIS_screen("REPT", "ACTV")
		CALL write_value_and_transmit(worker_number, 21, 13)
		'Making sure we are at the beginning of REPT/ACTV
		DO
			PF7
			EMReadScreen page_one, 2, 3, 78
		LOOP UNTIL page_one = " 1"
		
		rept_row = 7
		DO
			EMReadScreen case_number, 8, rept_row, 12
			case_number = trim(case_number)
			IF case_number <> "" THEN 
				case_number_array = case_number_array & case_number & "~~~"
				rept_row = rept_row + 1
				IF rept_row = 19 THEN 
					rept_row = 7 
					PF8
				END IF
			ELSE
				EXIT DO
			END IF
		LOOP 

	ELSEIF run_mode = "Excel File" THEN 
		'Opening the Excel file
		
		DO
			'file_location = InputBox("Please enter the file location.")
			
			Set objExcel = CreateObject("Excel.Application")
			Set objWorkbook = objExcel.Workbooks.Open(BrowseForFile)
			objExcel.Visible = True
			objExcel.DisplayAlerts = True
			
			confirm_file = MsgBox("Is this the correct file? Press YES to continue. Press NO to try again. Press CANCEL to stop the script.", vbYesNoCancel)
			IF confirm_file = vbCancel THEN 
				objWorkbook.Close
				objExcel.Quit
				stopscript
			ELSEIF confirm_file = vbNo THEN 
				objWorkbook.Close
				objExcel.Quit
			END IF
		LOOP UNTIL confirm_file = vbYes
		
		'Gathering the information from the user about the fields in Excel to look for.
		DO
			err_msg = ""
			DIALOG memo_from_excel_dlg
				IF ButtonPressed = 0 THEN stopscript
				IF isnumeric(excel_col) = FALSE AND len(excel_col) > 2 THEN 
					err_msg = err_msg & vbCr & "* Please do not use such a large column. The script cannot handle it."
				ELSE
					IF (isnumeric(right(excel_col, 1)) = TRUE AND isnumeric(left(excel_col, 1)) = FALSE) OR (isnumeric(right(excel_col, 1)) = FALSE AND isnumeric(left(excel_col, 1)) = TRUE) THEN
						err_msg = err_msg & vbCr & "* Please use a valid Column indicator. " & excel_col & " contains BOTH a letter and a number."
					ELSE
						call convert_excel_letter_to_excel_number(excel_col) 
						IF isnumeric(excel_row) = false or isnumeric(end_row) = false THEN err_msg = err_msg & vbCr & "* Please enter the Excel rows as numeric characters."
						IF end_row = "" THEN err_msg = err_msg & vbCr & "* Please enter an end to the search. The script needs to know when to stop searching."
					END IF
				END IF
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL err_msg = ""		
		
		CALL check_for_MAXIS(false)
		'Generating a MEMO for each case.
		FOR i = excel_row TO end_row
			IF objExcel.Cells(i, excel_col).Value <> "" THEN 
				case_number_array = case_number_array & objExcel.Cells(i, excel_col).Value & "~~~"
			END IF
		NEXT		
	END IF

CALL check_for_MAXIS(false)

'The business of sending memos
case_number_array = trim(case_number_array)
case_number_array = split(case_number_array, "~~~")

privileged_array = ""

FOR EACH case_number IN case_number_array
	forms_to_arep = ""					'clearing variables otherwise script will try to put a X as variable will remain Y between loops
	forms_to_swkr = ""
	IF case_number <> "" THEN 
		CALL navigate_to_MAXIS_screen("SPEC", "MEMO")
		'Checking for privileged
		EMReadScreen privileged_case, 40, 24, 2
		IF InStr(privileged_case, "PRIVILEGED") <> 0 THEN 
			privileged_array = privileged_array & case_number & "~~~"
		ELSE
			PF5
			'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
			row = 4                             'Defining row and col for the search feature.
			col = 1
			EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
			IF row > 4 THEN                     'If it isn't 4, that means it was found.
				arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
				call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
				EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
				call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
				PF5                                                     'PF5s again to initiate the new memo process
			END IF
			'Checking for SWKR
			row = 4                             'Defining row and col for the search feature.
			col = 1
			EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
			IF row > 4 THEN                     'If it isn't 4, that means it was found.
				swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
				call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
				EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
				call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
				PF5                                           'PF5s again to initiate the new memo process
			END IF
			EMWriteScreen "x", 5, 10                                        'Initiates new memo to client
			IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
			IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
			transmit    
			CALL write_variable_in_SPEC_MEMO(memo_text)
			PF4
		END IF
	END IF
NEXT

IF privileged_array <> "" THEN 
	privileged_array = replace(privileged_array, "~~~", vbCr)
	MsgBox "The script could not generate a memo for the following cases..." & vbCr & privileged_array
END IF

script_end_procedure("Success!!")
