all_scripts_array = ""

SET get_all_scripts = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a the URL for the text file

'Grabbing all the actions scripts
actions_url = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script%20Files/ACTIONS/ACTIONS%20-%20MAIN%20MENU.vbs"

'Grabbing all the bulk scripts
bulk_url = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script%20Files/BULK/BULK%20-%20MAIN%20MENU.vbs"

'grabbing all the Notes scripts
notes_url = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script%20Files/NOTES/NOTES%20-%20MAIN%20MENU.vbs"

'grabbing all the notices scripts
notices_url = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script%20Files/NOTICES/NOTICES%20-%20MAIN%20MENU.vbs"

all_url_array = actions_url & "UUDDLRLRBA" & bulk_url & "UUDDLRLRBA" & notes_array & "UUDDLRLRBA" & notices_array
all_url_array = split(all_url_array, "UUDDLRLRBA")

FOR EACH menu_url IN all_url_array
	msgbox menu_url
	get_all_scripts.open "GET", menu_url, FALSE						'Attempts to open the text file URL
	get_all_scripts.send													'Sends request
	IF get_all_scripts.Status = 200 THEN									'200 means great success
		Set fso = CreateObject("Scripting.FileSystemObject")				'Creates an FSO
		all_scripts = get_all_scripts.responseText								'Executes the script code
		all_scripts_array = all_scripts_array & script_array
	ELSE																	'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
		MsgBox 	"Something went wrong grabbing ACTIONS scripts."
		EXIT FOR
	END IF
NEXT	
