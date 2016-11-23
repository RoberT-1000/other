'Built by Robert Kalb and Charles Potter of Anoka County

'LOADING GLOBAL VARIABLES--------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Child Support\locally-installed-files\~globvar.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'Gathering stats
name_of_script = "ACTIONS - CHOOSE FAVORITE SCRIPTS.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
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
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'>>>>>>>>>>>>>>>>>>>>>>>>> SECTION 1 <<<<<<<<<<<<<<<<<<<<<<<<<<
'>>> The gobbins that happen before the user sees anything. <<<
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


'>>> Determining the location of the user's favorites list. 
'>>> This value should be stored in Global Variables for state-wide deployment.
network_location_of_favorites_text_file = "H:\my favorite cs scripts.txt"

'Creating the object to the URL a la text file
SET get_all_scripts = CreateObject("Msxml2.XMLHttp.6.0")

'switching up the script_repository because the all scripts file is not in the Script Files folder
all_scripts_repo = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/~complete-list-of-scripts.vbs"

'Building an array of all scripts
'Opening the URL for the given main menu
get_all_scripts.open "GET", all_scripts_repo, FALSE
get_all_scripts.send			
IF get_all_scripts.Status = 200 THEN	
	Set filescriptobject = CreateObject("Scripting.FileSystemObject")		
	Execute get_all_scripts.responseText
ELSE							
	'If the script cannot open the URL provided...
	MsgBox 	"Something went wrong with the URL: " & all_scripts_repo
	stopscript
END IF

'Warning/instruction box
MsgBox "This script will display a dialog with various scripts on it."  & vbNewLine &_
		"Any script you check will be added to your favorites menu.  " & vbNewline &_
		"Scripts you un-check will be removed. Once you are done " & vbNewLine &_
		"making your selection hit OK and your menu will be updated. " & vbNewLine & vbNewLine &_
		"- You will be unable to edit NEW Scripts and Recommended Scripts."

REDIM scripts_multidimensional_array(ubound(cs_scripts_array), 1)

'determining the number of each kind of script...by category
number_of_scripts = 0
actions_scripts = 0
bulk_scripts = 0
calc_scripts = 0
notes_scripts = 0
FOR i = 0 TO ubound(cs_scripts_array)
	'we are going to exclude navigation and utility scripts
	IF cs_scripts_array(i).category <> "nav" AND cs_scripts_array(i).category <> "utilities" THEN 
		number_of_scripts = number_of_scripts + 1
		IF cs_scripts_array(i).category = "actions" THEN 
			actions_scripts = actions_scripts + 1
		ELSEIF cs_scripts_array(i).category = "bulk" THEN 
			bulk_scripts = bulk_scripts + 1
		ELSEIF cs_scripts_array(i).category = "calculators" THEN 
			calc_scripts = calc_scripts + 1
		ELSEIF cs_scripts_array(i).category = "notes" THEN 
			notes_scripts = notes_scripts + 1
		END IF
	END IF
NEXT


'>>> If the user has already selected their favorites, the script will open that file and
'>>> and read it, storing the contents in the variable name ''user_scripts_array''
SET oTxtFile = (CreateObject("Scripting.FileSystemObject"))
With oTxtFile
	If .FileExists(network_location_of_favorites_text_file) Then
		Set fav_scripts = CreateObject("Scripting.FileSystemObject")
		Set fav_scripts_command = fav_scripts.OpenTextFile(network_location_of_favorites_text_file)
		fav_scripts_array = fav_scripts_command.ReadAll
		IF fav_scripts_array <> "" THEN user_scripts_array = fav_scripts_array
		fav_scripts_command.Close
	END IF
END WITH

'>>> Determining the width of the dialog from the number of scripts that are available...
'the dialog starts with a width of 800
dia_width = 800
'if a second column of actions scripts is needed, the dialog increases in width by 195
IF actions_scripts >= 40 AND actions_scripts <= 79 THEN 
	dia_width = dia_width + 195
	'if a third column of actions scripts is needed, the dialog increases in width by 195
ELSEIF actions_scripts >= 80 THEN 
	dia_width = dia_width + 195
END IF
'if a second column of bulk scripts is needed, the dialog increases in width by 195
IF bulk_scripts >= 40 AND bulk_scripts <= 79 THEN 
	dia_width = dia_width + 195
	'if a third column of bulk scripts is needed, the dialog increases in width by 195
ELSEIF bulk_scripts >= 80 THEN 
	dia_width = dia_width + 195
END IF
'if a second column of calc scripts is needed, the dialog increases in width by 195
IF calc_scripts >= 40 AND calc_scripts <= 79 THEN 
	dia_width = dia_width + 195
	'if a third column of calc scripts is needed, the dialog increases in width by 195
ELSEIF calc_scripts >= 80 THEN 
	dia_width = dia_width + 195
END IF
'if a second column of notes scripts is needed, the dialog increases in width by 195
IF notes_scripts >= 40 AND notes_scripts <= 79 THEN 
	dia_width = dia_width + 195
	'if a third column of notes scripts is needed, the dialog increases in width by 195
ELSEIF notes_scripts >= 80 AND notes_scripts <= 119 THEN 
	dia_width = dia_width + 195
	'if a fourth column of notes scripts is needed, the dialog increases in width by 195
ELSEIF notes_scripts >= 120 THEN 
	dia_width = dia_width + 195
END IF

'>>> Building the dialog
BeginDialog fav_dlg, 0, 0, dia_width, 440, "Select your favorites"
	ButtonGroup ButtonPressed
		OkButton 5, 5, 50, 15 
		CancelButton 55, 5, 50, 15
		PushButton 165, 5, 70, 15, "Reset Favorites", reset_favorites_button
	'>>> Creating the display of all scripts for selection (in checkbox form)
	script_position = 0		' <<< This value is tied to the number_of_scripts variable
		col = 10
		row = 30
	FOR i = 0 to ubound(cs_scripts_array)
		IF cs_scripts_array(i).category = "actions" THEN 
			'>>> Determining the positioning of the checkboxes.
			'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.	
			IF row = 430 THEN 
				row = 30
				col = col + 195
			END IF
			'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
			IF InStr(UCASE(replace(user_scripts_array, "-", " ")), UCASE(replace(cs_scripts_array(i).script_name, "-", " "))) <> 0 THEN  
				scripts_multidimensional_array(script_position, 1) = 1
			ELSE
				scripts_multidimensional_array(script_position, 1) = 0
			END IF
			scripts_multidimensional_array(script_position, 0) = "ACTIONS: " & replace(cs_scripts_array(i).file_name, ".vbs", "")
			CheckBox col, row, 185, 10, UCASE(replace(scripts_multidimensional_array(script_position, 0), "-", " ")), scripts_multidimensional_array(script_position, 1) 
			row = row + 10
			script_position = script_position + 1
		END IF
	NEXT
		col = col + 195
		row = 30
	FOR i = 0 to ubound(cs_scripts_array)
		IF cs_scripts_array(i).category = "bulk" THEN 
			'>>> Determining the positioning of the checkboxes.
			'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
			IF row = 430 THEN 
				row = 30
				col = col + 195
			END IF
			'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
			IF InStr(UCASE(replace(user_scripts_array, "-", " ")), UCASE(replace(cs_scripts_array(i).script_name, "-", " "))) <> 0 THEN  
				scripts_multidimensional_array(script_position, 1) = 1
			ELSE
				scripts_multidimensional_array(script_position, 1) = 0
			END IF
			scripts_multidimensional_array(script_position, 0) = "BULK: " & replace(cs_scripts_array(i).file_name, ".vbs", "")
			CheckBox col, row, 185, 10, UCASE(replace(scripts_multidimensional_array(script_position, 0), "-", " ")), scripts_multidimensional_array(script_position, 1) 
			row = row + 10
			script_position = script_position + 1
		END IF
	NEXT
		row = 30
		col = col + 195
	FOR i = 0 to ubound(cs_scripts_array)
		IF cs_scripts_array(i).category = "calculators" THEN 
			'>>> Determining the positioning of the checkboxes.
			'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
			IF row = 430 THEN 
				row = 30
				col = col + 195
			END IF
			'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
			IF InStr(UCASE(replace(user_scripts_array, "-", " ")), UCASE(replace(cs_scripts_array(i).script_name, "-", " "))) <> 0 THEN  
				scripts_multidimensional_array(script_position, 1) = 1
			ELSE
				scripts_multidimensional_array(script_position, 1) = 0
			END IF
			scripts_multidimensional_array(script_position, 0) = "CALC: " & replace(cs_scripts_array(i).file_name, ".vbs", "")
			CheckBox col, row, 185, 10, UCASE(replace(scripts_multidimensional_array(script_position, 0), "-", " ")), scripts_multidimensional_array(script_position, 1) 
			row = row + 10
			script_position = script_position + 1
		END IF
	NEXT
		col = col + 195
		row = 30
	FOR i = 0 to ubound(cs_scripts_array)
		IF cs_scripts_array(i).category = "notes" THEN 
			'>>> Determining the positioning of the checkboxes.
			'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
			IF row = 430 THEN 
				row = 30
				col = col + 195
			END IF
			'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
			IF InStr(UCASE(replace(user_scripts_array, "-", " ")), UCASE(replace(cs_scripts_array(i).script_name, "-", " "))) <> 0 THEN  
				scripts_multidimensional_array(script_position, 1) = 1
			ELSE
				scripts_multidimensional_array(script_position, 1) = 0
			END IF
			scripts_multidimensional_array(script_position, 0) = "NOTES - " & replace(cs_scripts_array(i).file_name, ".vbs", "")
			CheckBox col, row, 185, 10, "NOTES: " & UCASE(replace(scripts_multidimensional_array(script_position, 0), "-", " ")), scripts_multidimensional_array(script_position, 1) 
			row = row + 10
			script_position = script_position + 1
		END IF
	NEXT	
EndDialog

'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<
'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<
'>>>>>>>>>>>>>>>>>>>> SECTION 2 <<<<<<<<<<<<<<<<<<<<<
'>>> The gobbins that the user sees and makes do. <<<
'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<
'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<

DO
	DO
		'>>> Running the dialog
		Dialog fav_dlg
			'>>> Cancel confirmation
			IF ButtonPressed = 0 THEN 
				confirm_cancel = MsgBox("Are you sure you want to cancel? Press YES to cancel the script. Press NO to return to the script.", vbYesNo)
				IF confirm_cancel = vbYes THEN script_end_procedure("Script cancelled.")
			END IF
			'>>> If the user selects to reset their favorites selections, the script
			'>>> will go through the multi-dimensional array and reset all the values
			'>>> for position 1, thereby clearing the favorites from the display.
			IF ButtonPressed = reset_favorites_button THEN 
				FOR i = 0 to number_of_scripts
					scripts_multidimensional_array(i, 1) = 0
				NEXT
			END IF
	'>>> The exit condition for the first do/loop is the user pressing 'OK'
	LOOP UNTIL ButtonPressed <> 0 AND ButtonPressed <> reset_favorites_button
	'>>> Validating that the user does not select more than a prescribed number of scripts.
	'>>> Exceeding the limit will cause an exception access violation for the Favorites script when it runs.
	'>>> Currently, that value is 30. That is lower than previous because of the larger number of new scripts. (-Robert, 04/20/2016)
	double_check_array = ""
	FOR i = 0 to number_of_scripts
		IF scripts_multidimensional_array(i, 1) = 1 THEN double_check_array = double_check_array & scripts_multidimensional_array(i, 0) & "~"
	NEXT
	double_check_array = split(double_check_array, "~")
	IF ubound(double_check_array) > 29 THEN MsgBox "Your favorites menu is too large. Please limit the number of favorites to no greater than 30."
	'>>> Exit condition is the user having fewer than 30 scripts in their favorites menu.
LOOP UNTIL ubound(double_check_array) <= 29

'>>> Getting ready to write the user's selection to a text file and save it on a prescribed location on the network.
'>>> Building the content of the text file.	
FOR i = 0 to number_of_scripts - 1
	IF scripts_multidimensional_array(i, 1) = 1 THEN favorite_scripts = favorite_scripts & scripts_multidimensional_array(i, 0) & "~~~"
NEXT

'>>> After the user selects their favorite scripts, we are going to write (or overwrite) the list of scripts 
'>>> stored at H:\my favorite scripts.txt.
IF favorite_scripts <> "" THEN 
	SET updated_fav_scripts_fso = CreateObject("Scripting.FileSystemObject")
	SET updated_fav_scripts_command = updated_fav_scripts_fso.CreateTextFile(network_location_of_favorites_text_file, 2)
	updated_fav_scripts_command.Write(favorite_scripts)
	updated_fav_scripts_command.Close
	script_end_procedure("Success!! Your Favorites Menu has been updated.")
ELSE
	'>>> OR...if the user has selected no scripts for their favorite, the file will be deleted to 
	'>>> prevent the Favorites Menu from erroring out.
	'>>> Experience with worker_signature automation tells us that if the text file is blank, the favorites menu doth not work.
	oTxtFile.DeleteFile(network_location_of_favorites_text_file)
	script_end_procedure("You have updated your Favorites Menu, but you haven't selected any scripts. The next time you use the Favorites scripts, you will need to select your favorites.")
END IF

