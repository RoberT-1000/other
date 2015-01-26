'This script was developed by Charles Potter & Robert Kalb from Anoka County

'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - MA-EPD No Initial Premium"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\public assistance script files\script files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

EMConnect ""

call navigate_to_screen("spec", "wcom")
'pulled from previous menu selection
EMWriteScreen Approval_month, 3, 46
EMWriteScreen approval_year, 3, 51
EMWriteScreen "Y", 3, 74 'selects HC only
transmit

'array created in previous menu selection
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
      		EMWriteScreen "You are denied eligibility under Medical Assistance for", 3, 15
	      	EMWriteScreen "Employed Persons with Disabilities (MA-EPD) program because", 4, 15
	      	EMWriteScreen "the required premium was not paid by the due date, You may", 5, 15
			EMWriteScreen "request 'Good Cause' for late premium payment. This must be", 6, 15
			EMWriteScreen "approved by the Department of Human Services (DHS). To ", 7, 15
			EMWriteScreen "claim Good Cause, send a letter with your name, address,", 8, 15
			EMWriteScreen "case number and the reason for late payment to:", 9, 15
			EMWriteScreen "DHS MA-EPD Good Cause", 11, 15
			EMWriteScreen "P.O. Box 64967", 12, 15
			EMWriteScreen "St Paul, MN 55164-0967", 13, 15
			EMWriteScreen "Fax: 651 431 7563", 15, 15
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


If WCOM_count = 0 THEN
	MSGbox "No Waiting HC elig results were found in this month for this HH member."
	Stopscript	
ELSE
	MSGbox "Sucess! A WCOM has been added."
END IF


script_end_procedure("")