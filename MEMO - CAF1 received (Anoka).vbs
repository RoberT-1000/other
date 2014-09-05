'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - CAF1 received (Anoka)"
start_time = timer

''LOADING ROUTINE FUNCTIONS
'<<DELETE REDUNDANCIES!
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\County beta staging\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'SECTION 01
EMConnect ""
  row = 1
  col = 1
EMSearch "Case Nbr: ", row, col
EMReadScreen case_number, 8, row, col + 10
If case_number = "AR" then case_number = ""

BeginDialog CAF1_dialog, 5, 5, 176, 81, "CAF1 dialog"
  EditBox 85, 0, 50, 15, case_number
  EditBox 55, 20, 95, 15, CAF_date
  EditBox 80, 40, 85, 15, worker_sig
  ButtonGroup CAF1_dialog_ButtonPressed
    OkButton 35, 60, 50, 15
    CancelButton 90, 60, 50, 15
  Text 35, 5, 50, 10, "Case number:"
  Text 20, 25, 35, 10, "CAF date:"
  Text 5, 45, 65, 10, "Sign the case note:"
EndDialog

Do
  Dialog CAF1_dialog
  If CAF1_dialog_ButtonPressed = 0 then stopscript
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen MAXIS_check, 5, 1, 39
  IF MAXIS_check <> "MAXIS" then MsgBox "You need to be in MAXIS for this to work. Please try again."
  If case_number = "" or worker_sig = "" then MsgBox "You must fill in a case number and a signature before continuing."
  CAF_date = replace(CAF_date, ".", "/")
  If isdate(CAF_date) = False then Msgbox "You did not enter a valid date (MM/DD/YYYY format). Try again."
  If isdate(CAF_date) = True then 
    CAF_date = cdate(CAF_date)
    last_contact_day = CAF_date + 31
  End if
Loop until MAXIS_check = "MAXIS" and (case_number <> "" and isdate(CAF_date) = True and worker_sig <> "")

'SECTION 02
'This Do...loop gets back to SELF
do
  EMSendKey "<PF3>"
  EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"
EMWaitReady 1, 1
EMSetCursor 16, 43
EMSendKey "spec"
EMSetCursor 18, 43
EMSendkey "<eraseeof>" + case_number
EMSetCursor 21, 70
EMSendkey "memo" + "<enter>"
EMWaitReady 1, 1
'--------------ERROR PROOFING--------------
EMReadScreen still_self, 27, 2, 28 'This checks to make sure we've moved passed SELF.
If still_self = "Select Function Menu (SELF)" then StopScript 
EMReadScreen county, 4, 20, 14 'This will check the county. If this case is not x102, the script will stop.
If county <> "X102" then MsgBox "This case is not in Anoka County. Check your case number and try again."
If county <> "X102" then StopScript
'--------------END ERROR PROOFING--------------
EMSendKey "<PF5>"
EMWaitReady 1, 1
EMReadScreen memo_display_check, 12, 2, 33
If memo_display_check = "Memo Display" then MsgBox "You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production."
If memo_display_check = "Memo Display" then stopscript
EMSetCursor 5, 10
EMSendKey "x" + "<enter>"
EMWaitReady 1, 1
EMSetCursor 3, 15
EMSendKey "You recently applied for assistance in Anoka County on " & CAF_date & ". An interview is required to process your application." + "<newline>" + "<newline>"
EMSendKey "You must come into our office Monday through Friday, between 8:30am and 11:00am. Our office is located at:" + "<newline>" + "   2100 3rd Ave, Suite 400" + "<newline>" + "   Anoka, MN 55303" + "<newline>" + "<newline>"
EMSendKey "If you cannot attend an interview because of a hardship, please call our office at (763)422-7246." + "<newline>" + "<newline>"
EMSendKey "If we do not hear from you by " & last_contact_day & " we will deny your application."
EMSendKey "<PF4>"
EMWaitReady 1, 1

'SECTION 03
EMSetCursor 19, 22
EMSendKey "case"
EMSetCursor 19, 70
EMSendKey "note"
EMSendKey "<enter>"
EMWaitReady 1, 1
EMSendKey "<PF9>"
EMWaitReady 1, 1
EMSendKey "**CAF 1 received " & CAF_date & ", appt letter sent in MEMO**" + "<newline>"
EMSendKey "---" + "<newline>" + worker_sig

script_end_procedure("")