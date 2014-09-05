'This script was developed by Charles Potter & Robert Kalb from Anoka County

'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - Ended Six month spenddown"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\County beta staging\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script


EMConnect ""

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
EMSendKey "Your spenddown was decreased to zero due to new income standards. Providers that have already submitted claims to DHS will be reimbursed and in turn should either reimburse you or apply the amount to the current charges."
PF4

'Now, the case note
call navigate_to_screen("CASE","NOTE")
PF9
call write_new_line_in_case_note("**Ended Six Month Spenddown for " & first_name & " " & last_name & "**")
call write_new_line_in_case_note("Recalculated spenddown using the new 2014 income standards, enrollee does not have a spenddown. Updated MMIS.")
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)