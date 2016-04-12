'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
url = "https://raw.githubusercontent.com/theVKC/Anoka-PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
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
			StopScript
END IF


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
			EMReadScreen end_of_data, 11, 24, 2
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
		LOOP UNTIL row = 20
		PF8
		row = 7
	LOOP UNTIL end_of_data = "End of data"

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
BeginDialog inwd_dialog, 0, 0, 331, 280, "INWD Dialog"
  Text 10, 10, 40, 10, "Employer:"
  Text 75, 10, 85, 10, inwd_array(employer_number, 0)
  Text 10, 30, 75, 10, "Monthly Accrual"
  Text 20, 45, 50, 10, "Basic Support"
  Text 110, 45, 30, 10, inwd_array(employer_number, 1)
  Text 20, 60, 60, 10, "Spousal Maint."
  Text 110, 60, 30, 10, inwd_array(employer_number, 2)
  Text 20, 75, 60, 10, "Child Care"
  Text 110, 75, 30, 10, inwd_array(employer_number, 3)
  Text 20, 90, 60, 10, "Medical Support"
  Text 110, 90, 30, 10, inwd_array(employer_number, 4)
  Text 20, 105, 60, 10, "Other Support"
  Text 110, 105, 30, 10, inwd_array(employer_number, 5)
  Text 10, 125, 75, 10, "Monthly Accrual"
  Text 20, 140, 50, 10, "Basic Support"
  Text 110, 140, 30, 10, inwd_array(employer_number, 6)
  Text 20, 155, 60, 10, "Spousal Support"
  Text 110, 155, 30, 10, inwd_array(employer_number, 7)
  Text 20, 170, 60, 10, "Child Care"
  Text 110, 170, 30, 10, inwd_array(employer_number, 8)
  Text 20, 185, 60, 10, "Medical Support"
  Text 110, 185, 30, 10, inwd_array(employer_number, 9)
  Text 20, 200, 60, 10, "Other Support"
  Text 110, 200, 30, 10, inwd_array(employer_number, 10)
  Text 10, 220, 75, 10, "Additional 20%"
  Text 110, 220, 30, 10, inwd_array(employer_number, 11)
  Text 10, 235, 75, 10, "Total IW Amount"
  Text 110, 235, 30, 10, inwd_array(employer_number, 12)
  ButtonGroup ButtonPressed
    OkButton 115, 260, 50, 15
    CancelButton 165, 260, 50, 15

EndDialog

	DIALOG inwd_dialog
END FUNCTION


EMConnect ""
	'Dims the employer array to be used later
ReDim inwd_array(0, 12)
	'Creates the array of employers and values
CALL create_INWD_array
	'Outputs one dialog box per employer as defined by the number of employers
FOR i = 0 TO (UBound(inwd_array,1))
		'I = the array position of the current employer and passes this to the dialog box with the full array
		build_dialog(i,inwd_array)
NEXT
