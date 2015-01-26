'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - Open Period for Section 8"
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

'-----DIALOGS-----
BeginDialog section_8_dialog, 0, 0, 156, 255, "Section 8 HCV Program"
  EditBox 10, 210, 135, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 25, 235, 50, 15
    CancelButton 80, 235, 50, 15
  Text 10, 175, 140, 10, "Please sign your case note(s)."
  Text 5, 5, 145, 10, "This script will take the following actions:"
  Text 5, 25, 145, 20, "(1) Build a list of active MSA cases from your REPT/ACTV."
  Text 5, 50, 145, 35, "(2) Check STAT/PDED for ''Y'' on ''Shelter/Special Need.'' These cases have MSA Housing Assistance."
  Text 5, 90, 145, 40, "(3) Send a SPEC/MEMO to those clients, their AREPs and their Social Workers, notifying them of the open application period for the Section 8 Housing Choice Voucher program."
  Text 5, 135, 145, 10, "(4) Case note that the MEMO has been sent."
  Text 5, 150, 145, 20, "(5) Create a list in Microsoft Word of those cases that are receiving a MEMO."
  Text 10, 190, 140, 10, "Press OK to continue."
EndDialog

'-----THE SCRIPT-----
EMConnect ""
MAXIS_check_function

DO
	DIALOG section_8_dialog
		IF ButtonPressed = 0 THEN stopscript
		IF worker_signature = "" THEN MsgBox "Please sign the case note(s)."
LOOP UNTIL worker_signature <> "" AND ButtonPressed = -1

'Building the array to collect the worker's entire active MSA cases
CALL navigate_to_screen("REPT", "ACTV")
DO
	rept_actv_row = 7
	DO
		EMReadScreen last_page, 21, 24, 2
		EMReadScreen case_number, 8, rept_actv_row, 12
		case_number = replace(case_number, " ", "")
		EMReadScreen cash_active, 1, rept_actv_row, 54
		EMReadScreen cash_prog, 2, rept_actv_row, 51
		IF case_number <> "" AND cash_active = "A" AND cash_prog = "MS" THEN
			case_array = case_array & case_number & " "
		END IF
		rept_actv_row = rept_actv_row + 1
	LOOP UNTIL rept_actv_row = 19
	PF8
LOOP UNTIL last_page = "THIS IS THE LAST PAGE"

case_array = trim(case_array)
case_array = split(case_array, " ")

FOR EACH case_number IN case_array
	'-----This DO/LOOP is checking cases that are stuck in background. The script will try to navigate to STAT/PDED until it successfully gets there.------
	DO
		CALL navigate_to_screen("STAT", "PDED")
		ERRR_screen_check
		EMReadScreen at_pded, 4, 2, 50
	LOOP UNTIL at_pded = "PDED"

	pded_row = 5
	DO
		EMReadScreen pded_memb, 2, pded_row, 3
		pded_memb = replace(pded_memb, " ", "")
		IF pded_memb <> "" THEN 
			EMWriteScreen pded_memb, 20, 76
			transmit
		
			EMReadScreen special_need, 1, 18, 78
			pded_row = pded_row + 1
		END IF
	LOOP UNTIL pded_memb = "" OR special_need = "Y"

	IF special_need = "Y" THEN
		CALL navigate_to_screen("SPEC", "MEMO")
		PF5
		
		spec_memo_row = 5
		DO
			EMReadScreen memo_recip, 7, spec_memo_row, 12
			memo_recip = replace(memo_recip, " ", "")
			IF memo_recip = "CLIENT" OR memo_recip = "ALTREP" OR memo_recip = "SOCWKR" THEN	EMWriteScreen "X", spec_memo_row, 10
			spec_memo_row = spec_memo_row + 1
		LOOP UNTIL memo_recip = ""
		
		transmit
		
		EMWriteScreen "************************************************************", 3, 15
		EMWriteScreen "The Metro HRA will be accepting applications for Section 8", 4, 15
		EMWriteScreen "Housing Choice Vouchers during the following period:", 5, 15
		EMWriteScreen "     Tuesday, February 24, 2015 at 8:00 AM, through", 6, 15
		EMWriteScreen "     Friday, February 27, 2015 at 12:00 PM", 7, 15
		EMWriteScreen "************************************************************", 8, 15
		EMWriteScreen "Applications can be submitted online through the following", 9, 15
		EMWriteScreen "website:", 10, 15
		EMWriteScreen "     www.waitlistcheck.com/MN2707", 11, 15
		EMWriteScreen "************************************************************", 12, 15
		EMWriteScreen "Applications will not be available in HRA offices, nor can", 13, 15
		EMWriteScreen "they be emailed or faxed to the Metro HRA. Applicants", 14, 15
		EMWriteScreen "needing a reasonable accommodation may submit a written", 15, 15
		EMWriteScreen "request to the HRA no later than:", 16, 15
		EMWriteScreen "     Friday, February 13, 2015 at 4:30 PM.", 17, 15
		PF8
		EMWriteScreen "************************************************************", 3, 15
		EMWriteScreen "The Metro HRA's physical address is:", 4, 15
		EMWriteScreen "     390 Robert St N", 5, 15
		EMWriteScreen "     St. Paul, MN 55101", 6, 15
		EMWriteScreen "The Metro HRA's phone number is: (651)602-1428", 7, 15
		EMWriteScreen "************************************************************", 8, 15
		EMWriteScreen "You are receiving this notice because you are receiving MSA", 9, 15
		EMWriteScreen "and applying for housing assistance is a requirement of the", 10, 15
		EMWriteScreen "program.", 11, 15
		EMWriteScreen "************************************************************", 12, 15
		
	
		
		PF4
		
		CALL navigate_to_screen("CASE", "NOTE")
		PF9
		
		CALL write_variable_in_case_note("***SPEC/MEMO SENT TO CL, RE: Section 8 HCV Program***")
		CALL write_variable_in_case_note("* Notified CL of Metro HRA Housing Choice Voucher open application period.")
		CALL write_variable_in_case_note("---")
		CALL write_variable_in_case_note(worker_signature)
		PF3
		back_to_SELF		
		memo_array = memo_array & case_number & " "
	END IF
NEXT

memo_array = trim(memo_array)
memo_array = split(memo_array, " ")

'The script now creates a Word document showing the worker a list of cases that have received a SPEC/MEMO
Set objWord = CreateObject("Word.Application")
objWord.Visible = true
set objDoc = objWord.Documents.add()
Set objSelection = objWord.Selection
objselection.typetext "The following cases have been sent a SPEC/MEMO: "
objselection.TypeParagraph()
objselection.TypeParagraph()

FOR EACH memo_case IN memo_array
	objSelection.typetext "     * " & memo_case
	objselection.TypeParagraph()
NEXT

script_end_procedure("The script is finished running.")