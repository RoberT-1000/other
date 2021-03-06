'LOADING GLOBAL VARIABLES--------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Child Support\locally-installed-files\~globvar.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

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

'>>> Location of select favorites script on the network
network_location_of_select_favorites_script = "Q:\Blue Zone Scripts\Child Support\Script Files\County Customized\ANOKA - SELECT FAVORITE SCRIPTS.vbs"

'>>> Our script arrays. 
'>>> all_scripts_array will be built from the contents of the user's text file
'>>> new_scripts will be build automatically by looking at the description of each script in GitHub. If the description includes "NEW" then it is added to the array.
'>>> mandatory_array is pre-determined
all_scripts_array = ""
new_scripts = ""
mandatory_array = "ACTIONS - NCP LOCATE~~~ACTIONS - RECORD IW INFO~~~ACTIONS - SEND F0104 DORD MEMO~~~NOTES - ADJUSTMENTS~~~NOTES - ARREARS MANAGEMENT REVIEW~~~NOTES - CLIENT CONTACT~~~"

'>>> Creating the object needed to connect to the interwebs.
SET get_all_scripts = CreateObject("Msxml2.XMLHttp.6.0")
all_scripts_repo = script_repository & "~complete-list-of-scripts.vbs"
get_all_scripts.open "GET", all_scripts_repo, FALSE
get_all_scripts.send			
IF get_all_scripts.Status = 200 THEN	
	Set filescriptobject = CreateObject("Scripting.FileSystemObject")		
	Execute get_all_scripts.responseText
ELSE							
	'>>> Displaying the error message when the script fails to connect to a specific main menu.
	'>>> the replace & right bits are there to display the main menu in a way that is clear to the user.
	'>>> We are going to display the right length minus 99 because there are 99 characters between the start of the https and the last / before the main menu name.
	'>>> That length needs to be updated when we go state-wide.
	MsgBox("Something went wrong grabbing trying to locate All Scripts File. Please contact scripts administrator.")
	stopscript
END IF

'>>> Building the array of new scripts
'>>> If the description of the script includes the word "NEW" then the script file name is added to the array.
num_of_new_scripts = 0
new_array = ""
FOR i = 0 TO Ubound(cs_scripts_array)
	IF UCASE(cs_scripts_array(i).category) <> "NAV" AND _
		UCASE(cs_scripts_array(i).category) <> "UTILITIES" AND _
		DateDiff("D", cs_scripts_array(i).release_date, date) < 90 THEN
			new_array = new_array & UCASE(cs_scripts_array(i).category) & " - " & UCASE(replace(replace(cs_scripts_array(i).file_name, ".vbs", " "), "-", " ")) & "~~~"
	END IF
NEXT

'>>> Removing .vbs from the array for the prettification of the display to the users.
new_array = replace(new_array, ".vbs", "")

'>>> Custom function that builds the Favorites Main Menu dialog.
'>>> the array of the user's scripts
FUNCTION favorite_menu(user_scripts_array, mandatory_array, new_array, script_location, worker_signature)
	'>>> Splitting the array of all scripts. This is found on GitHub under Anoka-Specific Scripts
	user_scripts_array = trim(user_scripts_array)
	user_scripts_array = split(user_scripts_array, "~~~")

	mandatory_array = trim(mandatory_array)
	mandatory_array = split(mandatory_array, "~~~")
	
	new_array = trim(new_array)
	new_array = split(new_array, "~~~")
	
	num_of_user_scripts = ubound(user_scripts_array)
	num_of_mandatory_scripts = ubound(mandatory_array)
	num_of_new_scripts = ubound(new_array)
	
	num_of_scripts = num_of_user_scripts + num_of_mandatory_scripts + num_of_new_scripts
	
	ReDim all_scripts_array(num_of_scripts, 5)
	'position 0 = script name
	'position 1 = script directory
	'position 2 = button
	'position 3 = category
	'position 4 = script name without category
	'position 5 = state-supported true/false

	scripts_pos = 0
	FOR EACH script_name IN user_scripts_array
		IF script_name <> "" THEN 
			all_scripts_array(scripts_pos, 0) = script_name
			'>>> Creating the correct URL for the github call
			'>>> When we clean up this for state-wide deployment, we will need determine the appropriate network location for the agency custom scripts			
			IF left(script_name, 5) = "ANOKA" THEN 
				all_scripts_array(scripts_pos, 1) = "Q:\Blue Zone Scripts\Child Support\Script Files\County Customized\" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "ANOKA"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 7)
				all_scripts_array(scripts_pos, 5) = false
			ELSEIF left(script_name, 5) = "NOTES" THEN 
				all_scripts_array(scripts_pos, 1) = "/NOTES/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "NOTES"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 7)
				all_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_name, 7) = "ACTIONS" THEN 
				all_scripts_array(scripts_pos, 1) = "/ACTIONS/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "ACTIONS"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 9)
				all_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_name, 4) = "BULK" THEN 
				all_scripts_array(scripts_pos, 1) = "/BULK/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "BULK"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 6)
				all_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_name, 4) = "CALC" THEN 
				all_scripts_array(scripts_pos, 1) = "/CALCULATORS/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "CALCULATORS"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 6)
				all_scripts_array(scripts_pos, 5) = true
			END IF
			scripts_pos = scripts_pos + 1
		END IF	
	NEXT
	
	FOR EACH script_name IN mandatory_array
		IF script_name <> "" THEN 
			all_scripts_array(scripts_pos, 0) = script_name
			'>>> Creating the correct URL for the github call
			'>>> When we clean up this for state-wide deployment, we will need determine the appropriate network location for the agency custom scripts
			IF left(script_name, 5) = "ANOKA" THEN 
				all_scripts_array(scripts_pos, 1) = "Q:\Blue Zone Scripts\Child Support\Script Files\County Customized\" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "ANOKA"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 7)
				all_scripts_array(scripts_pos, 5) = false
			ELSEIF left(script_name, 5) = "NOTES" THEN 
				all_scripts_array(scripts_pos, 1) = "/NOTES/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "NOTES"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 7)
				all_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_name, 7) = "ACTIONS" THEN 
				all_scripts_array(scripts_pos, 1) = "/ACTIONS/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "ACTIONS"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 9)
				all_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_name, 4) = "BULK" THEN 
				all_scripts_array(scripts_pos, 1) = "/BULK/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "BULK"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 6)
				all_scripts_array(scripts_pos, 5) = true
			END IF
			scripts_pos = scripts_pos + 1
		END IF	
	NEXT

	FOR EACH script_name IN new_array
		IF script_name <> "" THEN 
			all_scripts_array(scripts_pos, 0) = script_name
			'>>> Creating the correct URL for the github call
			'>>> When we clean up this for state-wide deployment, we will need determine the appropriate network location for the agency custom scripts
			IF left(script_name, 5) = "ANOKA" THEN 
				all_scripts_array(scripts_pos, 1) = "Q:\Blue Zone Scripts\Child Support\Script Files\County Customized\" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "ANOKA"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 7)
				all_scripts_array(scripts_pos, 5) = false
			ELSEIF left(script_name, 5) = "NOTES" THEN 
				all_scripts_array(scripts_pos, 1) = "/NOTES/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "NOTES"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 7)
				all_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_name, 7) = "ACTIONS" THEN 
				all_scripts_array(scripts_pos, 1) = "/ACTIONS/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "ACTIONS"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 9)
				all_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_name, 4) = "BULK" THEN 
				all_scripts_array(scripts_pos, 1) = "/BULK/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "BULK"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 6)
				all_scripts_array(scripts_pos, 5) = true
			END IF
			scripts_pos = scripts_pos + 1
		END IF	
	NEXT	
	
	'>>> Determining the height parameters to enable the group boxes.
	actions_count = 0
	bulk_count = 0
	calc_count = 0
	notes_count = 0
	FOR i = 0 TO (ubound(user_scripts_array) - 1)
		IF all_scripts_array(i, 3) = "ACTIONS" THEN 
			actions_count = actions_count + 1
		ELSEIF all_scripts_array(i, 3) = "BULK" THEN 
			bulk_count = bulk_count + 1
		ELSEIF all_scripts_array(i, 3) = "CALCULATORS" THEN 
			calc_count = calc_count + 1
		ELSEIF all_scripts_array(i, 3) = "NOTES" THEN 
			notes_count = notes_count + 1
		END IF				
	NEXT
	
	'>>> Determining the height of the dialog.
	'>>> Each groupbox will require a minimum of 25 pixels. That is the height of the groupbox with 1 script PushButton
	'>>> The groupboxes need to grow 10 for each script pushbutton, so the dialog also needs to grow 10 for each script push button. However,
	'>>> 	the size of each groupbox will always be 15 plus (10 times the number of that kind of script)...
	dlg_height = 0
	IF actions_count <> 0 THEN dlg_height = 15 + (10 * actions_count)
	IF bulk_count <> 0 THEN dlg_height = dlg_height + 15 + (15 + (10 * bulk_count))
	IF calc_count <> 0 THEN dlg_height = dlg_height + 15 + (15 + (10 * bulk_count))
	IF notes_count <> 0 THEN dlg_height = dlg_height + 15 + (15 + (10 * notes_count))
	dlg_height = dlg_height + 5
	'>>> The dialog needs to be at least 185 pixels tall. If it is not...because the user has not selected a sufficient number of scripts...then
	'>>> the script needs to grow to 185.
	
	'>>> Adjusting the height if the user has fewer scripts than what is "recommended" plus the new scripts
	alt_dlg_height = 60 + (10 * (Ubound(mandatory_array) + 1)) + (10 * (Ubound(new_array) + 1))
	IF alt_dlg_height > dlg_height THEN dlg_height = alt_dlg_height
	
	'>>> Determining the start row for the push buttons
	'>>> The position of one groupbox will be determined from the existence of other groupboxes earlier in the alphabet.
	'>>> The actions start row is 10, and the end row will be 10 plus 15 (for the default height of the groupbox) plus 10 for each ACTIONS script
	IF actions_count <> 0 THEN 
		actions_start_row = 10
		actions_end_row = 10 + (15 + (10 * actions_count))
	ELSE
		'>>> ...or they will both be 0 when there are not ACTIONS scripts in the user's favorites.
		actions_start_row = 0
		actions_end_row = 0
	END IF
	'>>> The BULK groupbox start row will be determined by the end of the ACTIONS row...and so on.
	IF bulk_count <> 0 THEN 
		bulk_start_row = 10 + actions_end_row 
		bulk_end_row = bulk_start_row + (15 + (10 * bulk_count))
	ELSE
		bulk_start_row = actions_start_row
		bulk_end_row = actions_end_row			
	END IF
	IF calc_count <> 0 THEN 
		calc_start_row = 10 + bulk_end_row
		calc_end_row = calc_start_row + (15 + (10 * calc_count))
	ELSE
		calc_start_row = bulk_start_row
		calc_end_row = bulk_end_row
	END IF
	IF notes_count <> 0 THEN 
		notes_start_row = 10 + calc_end_row
		notes_end_row = notes_start_row + (15 + (10 * notes_count))
	ELSE
		notes_start_row = calc_start_row
		notes_end_row = calc_end_row
	END IF
	
	'>>> A nice decoration for the user. If they have used Update Worker Signature, then their signature is built into the dialog display.
	IF worker_signature <> "" THEN 
		dlg_name = worker_signature & "'s Favorite Scripts"
	ELSE
		dlg_name = "My Favorite Scripts"
	END IF
	
	'>>> The dialog
	BeginDialog favorites_dlg, 0, 0, 411, dlg_height, dlg_name & " "
  	  ButtonGroup ButtonPressed
		'>>> User's favorites
		'>>> Here, we are using the value for the script type start_row to determine the vertical position of each pushbutton.
		'>>> As we add a pushbutton, we need to increase the value for the start_row by 10 for that kind of script.
		FOR i = 0 TO (ubound(user_scripts_array) - 1)
			IF all_scripts_array(i, 3) = "ACTIONS" THEN 
				PushButton 20, actions_start_row + 10, 170, 10, UCASE(replace(all_scripts_array(i, 4), "-", " ")), all_scripts_array(i, 2)
				actions_start_row = actions_start_row + 10
			ELSEIF all_scripts_array(i, 3) = "BULK" THEN 
				PushButton 20, bulk_start_row + 10, 170, 10, UCASE(replace(all_scripts_array(i, 4), "-", " ")), all_scripts_array(i, 2)
				bulk_start_row = bulk_start_row + 10
			ELSEIF all_scripts_array(i, 3) = "CALCULATORS" THEN 
				PushButton 20, calc_start_row + 10, 170, 10, UCASE(replace(all_scripts_array(i, 4), "-", " ")), all_scripts_array(i, 2)
				calc_start_row = calc_start_row + 10
			ELSEIF all_scripts_array(i, 3) = "NOTES" THEN 
				PushButton 20, notes_start_row + 10, 170, 10, UCASE(replace(all_scripts_array(i, 4), "-", " ")), all_scripts_array(i, 2)
				notes_start_row = notes_start_row + 10			
			END IF
		NEXT

		'>>> Placing Mandatory Scripts
		FOR i = ubound(user_scripts_array) to (ubound(user_scripts_array) + (ubound(mandatory_array) - 1))
			right_hand_row = (20 + (10 * (i - num_of_user_scripts)))
			PushButton 220, right_hand_row, 180, 10, all_scripts_array(i, 0), all_scripts_array(i, 2)
		NEXT
		
		right_hand_row = right_hand_row + 30
		'>>> Placing new scripts
		FOR i = (ubound(user_scripts_array) + ubound(mandatory_array)) to (ubound(user_scripts_array) + ubound(mandatory_array) + (ubound(new_array) - 1))
			PushButton 220, right_hand_row, 180, 10, all_scripts_array(i, 0), all_scripts_array(i, 2)
			right_hand_row = right_hand_row + 10
		NEXT
		
		'>>> Placing groupboxes.
		'>>> All of the objects need to be placed at the end of the dialog. If they are not, it will throw off the positioning of the PushButtons
		'>>> which will, in turn, throw off the calculations for which script should be run.
		'>>> The height and position of each GroupBox is determed dynamically from the number of scripts in the groups previous.
		'>>> Mandatory and New are always going to be in the there, and located on the right hand side of the DLG.
        GroupBox 210, 10, 195, 5 + (10 * (Ubound(mandatory_array) + 1)), "Recommended Scripts"
		GroupBox 210, 20 + (10 * (Ubound(mandatory_array) + 1)), 195, 5 + (10 * (UBound(new_array) + 1)), "NEW SCRIPTS!!!"
		IF actions_count <> 0 THEN GroupBox 5, 10, 195, (15 + (10 * actions_count)), "ACTIONS"
		IF bulk_count <> 0 THEN GroupBox 5, actions_end_row + 10, 195, (15 + (10 * bulk_count)), "BULK"
		IF calc_count <> 0 THEN GroupBox 5, bulk_end_row + 10, 195, (15 + (10 * calc_count)), "CALCULATORS"
		IF notes_count <> 0 THEN GroupBox 5, calc_end_row + 10, 195, (15 + (10 * notes_count)), "NOTES"
		PushButton 210, dlg_height - 25, 70, 15, "Update Favorites", update_favorites_button
		CancelButton 355, dlg_height - 25, 50, 15
	EndDialog
	
	'>>> Loading the favorites dialog
	DIALOG favorites_dlg
		'>>> Cancelling the script if ButtonPressed = 0
		IF ButtonPressed = 0 THEN stopscript
		'>>> Giving user has the option of updating their favorites menu.
		'>>> We should try to incorporate the chainloading function of the new script_end_procedure to bring the user back to their favorites.
		IF buttonpressed = update_favorites_button THEN 
			call run_another_script(network_location_of_select_favorites_script)
			StopScript
		End if
		'>>> This tells the script which PushButton has been selected.
		'>>> We need to do ButtonPressed - 1 because of the way that the system assigns a value to ButtonPressed.
		'>>> When then favorites menu is launched from the Powerpad, the formula is ButtonPressed - 1. But if the menu is hidden behind another menu, then this formula is ButtonPressed - 1 - the number of other buttons ahead of the favorites menu button in that dialog tab order.
		script_location = all_scripts_array(ButtonPressed - 1, 1)  '!!!! THIS WILL NEED TO BE buttonpressed - (the number of objects created before the PushButtons...which is the dialog itself. don't move the order of the pushbuttons!!
		script_location = lcase(script_location)
		script_location = replace(script_location, " ", "-")
		script_location = replace(script_location, "bulk:-", "")
		script_location = replace(script_location, "actions:-", "")
		script_location = replace(script_location, "notes:-", "")
		script_location = replace(script_location, "notes---", "")
		script_location = replace(script_location, "actions---", "")
		script_location = replace(script_location, "bulk---", "")		
		script_location = replace(script_location, "calc:-", "")
		script_location = replace(script_location, "calc---", "")
END FUNCTION
'======================================

'The script starts HERE!!!-------------------------------------------------------------------------------------------------------------------------------------

'>>> The gobbins of the script that the user sees and makes do.
'>>> Declaring the text file storing the user's favorite scripts list.
Dim oTxtFile 
With (CreateObject("Scripting.FileSystemObject"))
	'>>> If the file exists, we will grab the list of the user's favorite scripts and run the favorites menu.
	If .FileExists("H:\my favorite cs scripts.txt") Then
		Set fav_scripts = CreateObject("Scripting.FileSystemObject")
		Set fav_scripts_command = fav_scripts.OpenTextFile("H:\my favorite cs scripts.txt")
		fav_scripts_array = fav_scripts_command.ReadAll
		IF fav_scripts_array <> "" THEN user_scripts_array = fav_scripts_array
		fav_scripts_command.Close
	ELSE
		'>>> ...otherwise, if the file does not exist, the script will require the user to select their favorite scripts.
		run_another_script(network_location_of_select_favorites_script)
	END IF
END WITH

'>>> Calling the function that builds the favorites menu.
CALL favorite_menu(user_scripts_array, mandatory_array, new_array, script_location, worker_signature)

script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master"

'>>> Running the script that is selected.
'>>> The first determination is whether the script is located on the agency's network.
IF left(script_location, 1) = "Q" THEN 
	'>>> Running the script if it is agency-custom script
	CALL run_another_script(script_location)
ELSE
	'>>> Running the script if it is stored in GitHub
	CALL run_from_GitHub(script_repository & script_location)
END IF
