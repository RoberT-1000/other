'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - MAIN MENU.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 THEN									'200 means great success'
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ''erroring out'' for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF

'CALL run_from_GitHub("https://raw.githubusercontent.com/RobertFewins-Kalb/Anoka-Specific-Scripts/master/REDIRECTION%20DOCS.vbs")

'stopscript

'DIALOGS---------------------------------------------------------------------------
BeginDialog ACTIONS_main_menu_dialog, 0, 0, 381, 240, "ACTIONS Main Menu"
  ButtonGroup ButtonPressed
    PushButton 5, 25, 85, 10, "Affidavit of Service Docs", ACTIONS_affidavit_of_service_button
    PushButton 5, 40, 60, 10, "DDPL Calculator", ACTIONS_DDPL_CALC_button
    PushButton 5, 55, 80, 10, "Estab NPA DORD Docs", ACTIONS_EST_DORD_NPA_button
    PushButton 5, 70, 80, 10, "Estab PA DORD Docs", ACTIONS_EST_DORD_PA_button
    PushButton 5, 85, 70, 10, "Find Name on CALI", ACTIONS_find_name_on_cali_button
    PushButton 5, 100, 30, 10, "Intake", ACTIONS_intake_button
    PushButton 5, 125, 60, 10, "PALC calculator", ACTIONS_PALC_calculator_button
    PushButton 5, 145, 60, 10, "Prorate Support", ACTIONS_prorate_support_button
    PushButton 5, 165, 65, 10, "Redirect Docs", ACTIONS_redirection_docs_button
    PushButton 5, 185, 75, 10, "Unreimb/Unins Docs", ACTIONS_un_un_button
    CancelButton 325, 220, 50, 15
    PushButton 300, 5, 75, 10, "PRISM Scripts in SIR", SIR_button
  Text 100, 25, 270, 10, "-- Sends Affidavits of Serivce to multiple participants on the case."
  Text 70, 40, 300, 10, "-- NEW 01/2016!! Calculates payments received during a specific date range."
  Text 90, 55, 280, 10, "-- NEW 01/2016!! Generates DORD docs for NPA case."
  Text 90, 70, 280, 10, "-- NEW 01/2016!! Generates DORD docs for PA case."
  Text 80, 85, 215, 10, "-- Searches CALI for a specific CP or NCP."
  Text 40, 100, 330, 15, "-- Creates various documents related to Child Support intake, as well as DORD documents, and enters a note on CAAD."
  Text 70, 125, 230, 10, "-- Calculates voluntary and involuntary amounts from the PALC screen."
  Text 70, 145, 225, 10, "- Calculator for deteremining pro-rated support for partial months."
  Text 75, 165, 290, 10, "-- Creates redirection docs and redirection worklist items."
  Text 85, 185, 290, 10, "-- Prints DORD docs for collecting unreimbursed and unisured expenses."
EndDialog



'THE SCRIPT-----------------------------------------------------------------------------------------------

'Shows the dialog
DO
	Dialog ACTIONS_main_menu_dialog
	If buttonpressed = cancel then stopscript
	IF ButtonPressed = SIR_button THEN CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/PRISMscripts/PRISM%20script%20wiki/Forms/AllPages.aspx")
LOOP UNTIL ButtonPressed <> SIR_button
IF ButtonPressed = ACTIONS_affidavit_of_service_button 		THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - AFFIDAVIT OF SERVICE BY MAIL DOCS.vbs")
IF ButtonPressed = ACTIONS_DDPL_CALC_button 				THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - DDPL CALCULATOR.vbs")
IF ButtonPressed = ACTIONS_EST_DORD_NPA_button 				THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - EST DORD NPA DOCS.vbs")
IF ButtonPressed = ACTIONS_EST_DORD_PA_button 				THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - ESTB DORD PA DOCS.vbs")
IF ButtonPressed = ACTIONS_find_name_on_cali_button 		THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - FIND NAME ON CALI.vbs")
IF ButtonPressed = ACTIONS_prorate_support_button 			THEN call run_from_GitHub(script_repository & "ACTIONS/ACTIONS - PRORATE SUPPORT.vbs")
IF ButtonPressed = ACTIONS_intake_button 					THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - INTAKE.vbs")
IF ButtonPressed = ACTIONS_PALC_calculator_button 			THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - PALC CALCULATOR.vbs")
IF ButtonPressed = ACTIONS_redirection_docs_button 			THEN 
	script_url = script_repository & "ACTIONS/ACTIONS - REDIRECT DOCS.vbs"
	script_url = replace(script_url, " ", "%20")
	CreateObject("WScript.Shell").Run(script_url)
'	msgbox true
	'CALL run_from_GitHub("https://raw.githubusercontent.com/RobertFewins-Kalb/Anoka-Specific-Scripts/master/REDIRECTION DOCS.vbs")
End IF
IF ButtonPressed = ACTIONS_un_un_button 					THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - UNREIMBURSED UNINSURED DOCS.vbs")
