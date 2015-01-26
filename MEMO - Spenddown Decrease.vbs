'This script was developed by Charles Potter & Robert Kalb from Anoka County

'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - Spenddown Decrease"
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


BeginDialog spenddown_decrease_dialog, 0, 0, 141, 165, "Spenddown Decrease"
  EditBox 10, 20, 65, 15, month1
  EditBox 85, 20, 50, 15, spenddown1
  EditBox 10, 40, 65, 15, month2
  EditBox 85, 40, 50, 15, spenddown2
  EditBox 10, 60, 65, 15, month3
  EditBox 85, 60, 50, 15, spenddown3
  EditBox 10, 80, 65, 15, month4
  EditBox 85, 80, 50, 15, spenddown4
  EditBox 10, 100, 65, 15, month5
  EditBox 85, 100, 50, 15, spenddown5
  EditBox 10, 120, 65, 15, month6
  EditBox 85, 120, 50, 15, spenddown6
  ButtonGroup ButtonPressed
    OkButton 20, 145, 50, 15
    CancelButton 75, 145, 50, 15
  Text 15, 10, 55, 10, "Affected Month"
  Text 95, 10, 50, 10, "Spenddown"
EndDialog

EMConnect ""

Dialog spenddown_decrease_dialog
  IF buttonpressed = 0 THEN stopscript

'memo_text = "Due to new income standards, your spenddown for the following months was decreased to:"
'IF month1 <> "" THEN memo_text = memo_text & month1 & " ($" & spenddown1 & ")"
'IF month2 <> "" THEN memo_text = memo_text & ", " & month2 & " ($" & spenddown2 & ")"
'IF month3 <> "" THEN memo_text = memo_text & ", " & month3 & " ($" & spenddown3 & ")"
'IF month4 <> "" THEN memo_text = memo_text & ", " & month4 & " ($" & spenddown4 & ")"
'IF month5 <> "" THEN memo_text = memo_text & ", " & month5 & " ($" & spenddown5 & ")"
'IF month6 <> "" THEN memo_text = memo_text & ", " & month6 & " ($" & spenddown6 & ")"

CALL navigate_to_screen("SPEC","MEMO")
PF5
EMReadScreen memo_display_check, 12, 2, 33
If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
EMWriteScreen "x", 5, 10
transmit

'Sends the home key to get to the top of the memo.
EMSendKey "<home>" 
EMSendKey "REGARDING: " & first_name & " " & last_name
EMSendKey "<newline>" & "<newline>"
EMSendKey "Due to new income standards, your spenddown for the following months was decreased to:"
EMSendKey "<newline>"
IF month1 <> "" THEN
  EMSendKey "     " & month1 & "      $" & spenddown1
  EMSendKey "<newline>"
END IF
IF month2 <> "" THEN
  EMSendKey "     " & month2 & "      $" & spenddown2
  EMSendKey "<newline>"
END IF
IF month3 <> "" THEN
  EMSendKey "     " & month3 & "      $" & spenddown3
  EMSendKey "<newline>"
END IF
IF month4 <> "" THEN
  EMSendKey "     " & month4 & "      $" & spenddown4
  EMSendKey "<newline>"
END IF
IF month5 <> "" THEN
  EMSendKey "     " & month5 & "      $" & spenddown5
  EMSendKey "<newline>"
END IF
IF month6 <> "" THEN
  EMSendKey "     " & month6 & "      $" & spenddown6
  EMSendKey "<newline>"
END IF
EMSendKey "Providers that have already submitted claims to DHS will be reimbursed and in turn should either reimburse you or apply the amount to current charges."
PF4

'Now, the case note
call navigate_to_screen("CASE","NOTE")
PF9
call write_new_line_in_case_note("**Decreased Monthly Spenddown for " & first_name & " " & last_name & "**")
call write_new_line_in_case_note("-Recalculated spenddown using the new 2014 income standards.")
call write_new_line_in_case_note("-For the following months, enrollee's spenddown decreased to:")
IF month1 <> "" THEN
  EMSendKey "     " & month1 & " ... $" & spenddown1
  EMSendKey "<newline>"
END IF
IF month2 <> "" THEN
  EMSendKey "     " & month2 & " ... $" & spenddown2
  EMSendKey "<newline>"
END IF
IF month3 <> "" THEN
  EMSendKey "     " & month3 & " ... $" & spenddown3
  EMSendKey "<newline>"
END IF
IF month4 <> "" THEN
  EMSendKey "     " & month4 & " ... $" & spenddown4
  EMSendKey "<newline>"
END IF
IF month5 <> "" THEN
  EMSendKey "     " & month5 & " ... $" & spenddown5
  EMSendKey "<newline>"
END IF
IF month6 <> "" THEN
  EMSendKey "     " & month6 & " ... $" & spenddown6
  EMSendKey "<newline>"
END IF
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)