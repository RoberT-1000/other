SET get_all_scripts = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a the URL for the text file

'Grabbing all the actions scripts
actions_url = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script%20Files/ACTIONS/ACTIONS%20-%20MAIN%20MENU.vbs"

'Grabbing all the bulk scripts
bulk_url = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script%20Files/BULK/BULK%20-%20MAIN%20MENU.vbs"

'grabbing all the Notes scripts
notes_url = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script%20Files/NOTES/NOTES%20-%20MAIN%20MENU.vbs"

'grabbing all the notices scripts
notices_url = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script%20Files/ACTIONS/ACTIONS%20-%20MAIN%20MENU.vbs"

get_all_scripts.open "GET", actions_url, FALSE						'Attempts to open the text file URL
get_all_scripts.send													'Sends request
IF get_all_scripts.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")				'Creates an FSO
	all_scripts = get_all_scripts.responseText								'Executes the script code
ELSE																	'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something went wrong grabbing ACTIONS scripts."
	stopscript
END IF

get_all_scripts.open "GET", bulk_url, FALSE						'Attempts to open the text file URL
get_all_scripts.send													'Sends request
IF get_all_scripts.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")				'Creates an FSO
	all_scripts = get_all_scripts.responseText								'Executes the script code
ELSE																	'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something went wrong grabbing BULK scripts."
	stopscript
END IF

get_all_scripts.open "GET", notes_url, FALSE						'Attempts to open the text file URL
get_all_scripts.send													'Sends request
IF get_all_scripts.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")				'Creates an FSO
	all_scripts = get_all_scripts.responseText								'Executes the script code
ELSE																	'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something went wrong grabbing NOTES scripts."
	stopscript
END IF

get_all_scripts.open "GET", notices_url, FALSE						'Attempts to open the text file URL
get_all_scripts.send													'Sends request
IF get_all_scripts.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")				'Creates an FSO
	all_scripts = get_all_scripts.responseText								'Executes the script code
ELSE																	'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something went wrong grabbing NOTICES scripts."
	stopscript
END IF
