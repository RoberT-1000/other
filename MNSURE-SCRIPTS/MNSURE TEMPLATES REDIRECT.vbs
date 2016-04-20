'LOADING GLOBAL VARIABLES--------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Public assistance script files\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
If beta_agency = "" or beta_agency = True then
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
Else
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
End if

SET req = CreateObject("Msxml2.XMLHttp.6.0") 'Creates an object to get a URL
req.open "GET", url, FALSE	'Attempts to open the URL
req.send 'Sends request

IF req.Status = 200 THEN	'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject") 'Creates an FSO
	Execute req.responseText 'Executes the script code
ELSE	'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox "Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
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

BeginDialog mnsure_case_notes_dlg, 0, 0, 331, 215, "MNSure Case Note Templates"
  ButtonGroup ButtonPressed
    PushButton 10, 20, 125, 15, "Address Change - Within County **", addr_change_in_county_button
    PushButton 10, 35, 125, 15, "Address Change - Out of County **", addr_change_out_of_county_button
    PushButton 10, 50, 125, 15, "AREP Task Completed **", arep_updated_button
    PushButton 10, 65, 125, 15, "Client Contact", client_contact_button
    PushButton 10, 80, 125, 15, "Document(s) Received **", docs_received_button
    PushButton 10, 95, 125, 15, "Request for Information Sent **", RFI_sent_button
    CancelButton 275, 195, 50, 15
  Text 10, 5, 325, 10, "Scripts followed by a double asterisk (**) indicate scripts that can case note and TIKL in MAXIS."
EndDialog



DIALOG mnsure_case_notes_dlg
	IF ButtonPressed = 0 THEN stopscript
	
IF ButtonPressed = addr_change_out_of_county_button THEN
	CALL run_another_script("Q:\Blue Zone Scripts\Public Assistance Script Files\Script Files\County Customized\MNSURE NOTES - OUT OF COUNTY ADDRESS CHANGE.vbs")
	stopscript
ELSEIF ButtonPressed = addr_change_in_county_button THEN 
	CALL run_another_script("Q:\Blue Zone Scripts\Public Assistance Script Files\Script Files\County Customized\MNSURE NOTES - WITHIN COUNTY ADDRESS CHANGE.vbs")
	stopscript
ELSEIF ButtonPressed = arep_updated_button THEN
	CALL run_another_script("Q:\Blue Zone Scripts\Public Assistance Script Files\Script Files\County Customized\MNSURE NOTES - AREP TASK.vbs")
	stopscript
ELSEIF ButtonPressed = docs_received_button THEN 
	CALL run_another_script("Q:\Blue Zone Scripts\Public Assistance Script Files\Script Files\County Customized\MNSURE NOTES - MNSURE DOCS RECEIVED.vbs")
	stopscript
ELSEIF ButtonPressed = client_contact_button THEN 
	CALL run_another_script("Q:\Blue Zone Scripts\Public Assistance Script Files\Script Files\County Customized\MNSURE NOTES - CLIENT CONTACT.vbs")
	stopscript
ELSEIF ButtonPressed = RFI_sent_button THEN
	CALL run_another_script("Q:\Blue Zone Scripts\Public Assistance Script Files\Script Files\County Customized\MNSURE NOTES - RFI SENT.vbs")
	stopscript
END IF
