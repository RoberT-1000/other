'This script was developed by Charles Potter & Robert Kalb from Anoka County

'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - MAGI Approved"
start_time = timer

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


EMConnect ""

call navigate_to_screen("spec", "wcom")
EMWriteScreen Approval_month, 3, 46
EMWriteScreen approval_year, 3, 51
EMWriteScreen "Y", 3, 74
transmit

FOR each HH_member in HH_member_array
	DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
		EMReadScreen more_pages, 8, 18, 72
		IF more_pages = "MORE:  -" THEN PF7
	LOOP until more_pages <> "MORE:  -"

	read_row = 7
	DO
		EMReadscreen reference_number, 2, read_row, 62 
		EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
		If waiting_check = "Waiting" and reference_number = HH_member THEN 'checking program type and if it's been printed, needs more fool proofing
			EMSetcursor read_row, 13
			EMSendKey "x"
			Transmit
			pf9
		      EMSetCursor 03, 15
      		EMWriteScreen "You will remain eligible for Medical Assistance because of", 3, 15
	      	EMWriteScreen "new rules and guidelines. (Authority: 42 C.F.R. 435.603(a)", 4, 15
	      	EMWriteScreen "(3); Section 1902(e)(14)(A)", 5, 15
		      PF4
			PF3
			WCOM_count = WCOM_count + 1
			exit do
		ELSE
			read_row = read_row + 1
		END IF
		IF read_row = 18 THEN
			PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18??
			read_row = 7
		End if
	LOOP until reference_number = "  "
NEXT


If WCOM_count <> 0 THEN
	back_to_self
	call navigate_to_screen("Case", "Note")	
	pf9
	call write_new_line_in_case_note("***Magi renewal***")
	FOR EACH HH_member IN HH_member_array
 	 magi_case_note_line_one = "* Member " & HH_member & " remains eligible for Medical Assistance for an additional year"
	  magi_case_note_line_two = "  because of new rules and guidelines."
	  call write_new_line_in_case_note(magi_case_note_line_one)
	  call write_new_line_in_case_note(magi_case_note_line_two)
	NEXT
	call write_new_line_in_case_note("---")
	call write_new_line_in_case_note(worker_signature)
ELSE
	MSGbox "No Waiting HC elig results were found in this month for this HH member."
	Stopscript	
END IF

script_end_procedure("")
