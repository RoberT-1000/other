'Master WCOM Created by Charles Potter and Robert Kalb from Anoka County

'Informational front-end message, date dependent.
If datediff("d", "06/23/2014", now) < 7 then MsgBox "This script has been added as of 06/23/2014! Here's what it does:" & chr(13) & chr(13) & "It will prompt you to enter the case number, renewal month information and then allow you to select which HH members need to be identified as MAGI renewals. Then it will add the text given by DHS to the wcom and to a seperate case note. if you have any questions please email Robert or Charles."

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - Master WCOM"
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


'Functions========================================================
function run_WCOM_script(WCOM_script_path)
  Set run_another_WCOM_fso = CreateObject("Scripting.FileSystemObject")
  Set fso_WCOM_command = run_another_WCOM_fso.OpenTextFile(WCOM_script_path)
  text_from_the_other_WCOM_script = fso_WCOM_command.ReadAll
  fso_WCOM_command.Close
  Execute text_from_the_other_WCOM_script
  stopscript
end function

'Dialogs===========================================================
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

BeginDialog wcom_menu_dialog, 0, 0, 141, 175, "WCOM menu"
  CheckBox 10, 25, 120, 10, "MAGI Remains Eligible", MAGI_renewal_check
  CheckBox 10, 40, 115, 10, "Decreased Spenddown", Decreased_spenddown_check
  CheckBox 10, 55, 110, 10, "Ended Monthly Spendown", Ended_monthly_spenddown_check
  CheckBox 10, 70, 110, 10, "Ended Six-Month Spenddown", Ended_six_month_spenddown_check
  CheckBox 10, 85, 105, 10, "New APP Denied", New_app_denied_check
  CheckBox 10, 100, 105, 10, "Enrollee Closed", Enrollee_closed_check
  CheckBox 10, 115, 115, 10, "New APP Eligible on Spenddown", New_app_elig_check
  CheckBox 10, 130, 120, 10, "MA-EPD No Initial Premium Paid", MAEPD_no_initial_premium_check
  ButtonGroup ButtonPressed
    OkButton 20, 150, 50, 15
    CancelButton 70, 150, 50, 15
  Text 10, 10, 120, 10, "Please select ONE situation..."
EndDialog


EMConnect ""

transmit

'Performs a MAXIS check-----------------------------------------------
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

DO
  dialog WCOM_menu_dialog
  IF buttonpressed = 0 then stopscript
  check_count = 0
  IF Magi_renewal_check = 1 THEN check_count = check_count + 1
  IF Decreased_spenddown_check = 1 THEN check_count = check_count + 1
  IF Ended_monthly_spenddown_check = 1 THEN check_count = check_count + 1
  IF Ended_six_month_spenddown_check = 1 THEN check_count = check_count + 1
  IF New_app_denied_check = 1 THEN check_count = check_count + 1
  IF Enrollee_closed_check = 1 THEN check_count = check_count + 1
  IF New_app_elig_check = 1 THEN check_count = check_count + 1
  IF MAEPD_no_initial_premium_check = 1 THEN check_count = check_count + 1
  IF check_count <> 1 THEN MSGBox "You may select ONE situation only."
LOOP UNTIL check_count = 1

DO
  person_count = 0
  HH_member_array = ""
  call navigate_to_screen("stat", "memb")
  EMReadScreen ERRR_check, 4, 2, 52		'Error prone case checking
  If ERRR_check = "ERRR" then transmit	'transmitting if case is error prone
  call HH_member_custom_dialog(HH_member_array)
  back_to_SELF
  FOR EACH person in HH_member_array
    person_count = person_count + 1
  NEXT  
  IF ((Decreased_spenddown_check = 1 or Ended_monthly_spenddown_check = 1 or Ended_six_month_spenddown_check = 1) AND person_count <> 1) THEN 
    MSGBox "The situation selected can only support one household member at a time. Please select just one household member and try again."
    back_to_SELF
  END IF
LOOP UNTIL ((Decreased_spenddown_check = 1 or Ended_monthly_spenddown_check = 1 or Ended_six_month_spenddown_check = 1) AND person_count = 1) OR (Magi_renewal_check = 1 or New_app_denied_check = 1 or Enrollee_closed_check = 1 or New_app_elig_check = 1 or MAEPD_no_initial_premium_check = 1)

call navigate_to_screen("STAT","MEMB")
FOR EACH person in HH_member_array
  EMWriteScreen person, 20, 76
  transmit
  EMReadScreen first_name, 12, 6, 63
    first_name = replace(first_name, "_", "")   
  EMReadScreen last_name, 25, 6, 30
    last_name = replace(last_name, "_", "")
NEXT

'Run other scripts depending on which script is selected=====================================
'Runs the MAGI WCOM script
IF Magi_renewal_check = 1 THEN run_WCOM_script("Q:\Blue Zone Scripts\public assistance script files\script files\County customized\MEMO - MAGI Approved.vbs")

'Runs the Decreased Spenddown script
IF Decreased_spenddown_check = 1 THEN run_WCOM_script("Q:\Blue Zone Scripts\public assistance script files\script files\County customized\MEMO - Spenddown Decrease.vbs")

'Runs the Ended Monthly Spenddown script
IF Ended_monthly_spenddown_check = 1 THEN run_WCOM_script("Q:\Blue Zone Scripts\public assistance script files\script files\County customized\MEMO - Ended Monthly Spenddown.vbs")

'Runs the Ended Six Month Spenddown script
IF Ended_six_month_spenddown_check = 1 THEN run_WCOM_script("Q:\Blue Zone Scripts\public assistance script files\script files\County customized\MEMO - Ended Six Month Spenddown.vbs")

'Runs the Closed/Denied script
IF (Enrollee_closed_check = 1 or New_app_denied_check = 1) THEN run_WCOM_script("Q:\Blue Zone Scripts\public assistance script files\script files\County customized\MEMO - App Denied Enrollee Closed.vbs")

'Runs the New APPL Elig script
IF New_app_elig_check = 1 THEN run_WCOM_script("Q:\Blue Zone Scripts\public assistance script files\script files\MEMO - New App Elig.vbs")

'Runs the MA-EPD No Initial Premium script
IF MAEPD_no_initial_premium_check = 1 THEN run_WCOM_script("Q:\Blue Zone Scripts\public assistance script files\script files\county customized\MEMO - MA-EPD No Initial Premium.vbs")


script_end_procedure("")
