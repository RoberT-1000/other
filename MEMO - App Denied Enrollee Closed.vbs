'This script was developed by Charles Potter & Robert Kalb from Anoka County

'GATHERING STATS----------------------------------------------------------------------------------------------------'
name_of_script = "MEMO - App Denied Enrollee Closed"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\County beta staging\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

EMConnect ""

call navigate_to_screen("SPEC","WCOM")
EMWriteScreen Approval_month, 3, 46
EMWriteScreen approval_year, 3, 51
EMWriteScreen "Y", 3, 74
transmit

WCOM_count = 0
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
                  EMWriteScreen "You do not qualify for Medical Assistance. We recommend", 3, 15
                  EMWriteScreen "you shop for health care coverage in MNSure as soon as", 4, 15
                  EMWriteScreen "possible. MNSure is a new online marketplace where", 5, 15
                  EMWriteScreen "Minnesotans can apply to get quality, affordable health", 6, 15
                  EMWriteScreen "care coverage. For more information, go to:", 7, 15
                  EMWriteScreen "     http://www.mnsure.org", 9, 15
                  EMWriteScreen "If you need help completing the online application or to", 11, 15
                  EMWriteScreen "request a paper application, call the MHCP Member Help", 12, 15
                  EMWriteScreen "Desk at 651-431-2670 or 1-800-657-3739. TTY users can call", 13, 15
                  EMWriteScreen "through Minnesota Relay at 711.", 14, 15
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

IF wcom_count <> 0 THEN 
  MSGBox "The script has successfully added a worker comment to " & WCOM_count & " notice(s). Please case note if you have not already done so."
ELSE
  MSGbox "No Waiting HC elig results were found in this month."
END IF