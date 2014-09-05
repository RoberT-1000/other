'MAGI renewal Created by Charles Potter and Robert Kalb from Anoka County

'Informational front-end message, date dependent.
If datediff("d", "06/23/2014", now) < 7 then MsgBox "This script has been added as of 06/23/2014! Here's what it does:" & chr(13) & chr(13) & "It will prompt you to enter the case number, renewal month information and then allow you to select which HH members need to be identified as MAGI renewals. Then it will add the text given by DHS to the wcom and to a seperate case note. if you have any questions please email Robert or Charles."

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - MAGI WCOM"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\County beta staging\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'Functions===========================================================
'Performs a MAXIS check-----------------------------------------------
BeginDialog magi_dlg, 0, 0, 196, 100, "MAGI WCOM"
  EditBox 70, 15, 60, 15, case_number
  EditBox 70, 35, 30, 15, Approval_month
  EditBox 160, 35, 30, 15, approval_year
  EditBox 80, 55, 60, 15, worker_signature
  Text 10, 20, 55, 10, "Case Number: "
  Text 10, 40, 55, 10, "Approval Month:"
  Text 105, 40, 55, 10, "Approval Year:"
  Text 10, 60, 70, 10, "Worker signature: "
  ButtonGroup ButtonPressed
    OkButton 50, 80, 50, 15
    CancelButton 105, 80, 50, 15
EndDialog

EMConnect ""

transmit

maxis_check_function

row = 1
col = 1
EMSearch "Case Nbr:", row, col
If row <> 0 then 
  EMReadScreen case_number, 8, row, col + 10
  case_number = replace(case_number, "_", "")
  case_number = trim(case_number)
End if

If isnumeric(case_number) = False then case_number = ""

DO
 DO
  Do
	dialog magi_dlg
	If buttonpressed = 0 then stopscript
      IF len(Approval_month) = 1 THEN Approval_month = "0" & Approval_month    'Converts the approval month to a 2 digit string to get MAXIS to behave appropriately
      IF len(approval_year) <> 2 THEN MSGBox("Approval Year must be last 2 digits of the year.")
	If worker_signature = "" then msgbox("Please sign your name")
      IF isnumeric(case_number) = FALSE THEN msgbox("You need a valid case number -- no letters or special characters.")
      IF len(case_number) > 8 THEN msgbox("You need a valid case number -- no longer than 8 digits.")
  LOOP UNTIL len(approval_year) = 2
 LOOP UNTIL case_number <> "" and isnumeric(case_number) = TRUE and len(case_number) < 9
Loop until worker_signature <> ""

call navigate_to_screen("stat", "memb")
EMReadScreen ERRR_check, 4, 2, 52		'Error prone case checking
If ERRR_check = "ERRR" then transmit	'transmitting if case is error prone
call HH_member_custom_dialog(HH_member_array)
back_to_SELF

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