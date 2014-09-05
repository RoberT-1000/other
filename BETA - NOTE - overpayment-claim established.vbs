'STATS GATHERING----------------------------------------------------------------------------------------------------
'name_of_script = "NOTE - overpayment-claim established"
'start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\County beta staging\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script


'SECTION 02: DIALOGS

county_code = "02"

If county_code = "02" then text_for_dialog = "Remember to ''staple'' the supporting documents to the claim form, and send to your supervisor for approval!"

BeginDialog overpayment_dialog, 0, 0, 266, 260, "Overpayment dialog"
  EditBox 60, 5, 70, 15, case_number
  EditBox 120, 25, 140, 15, programs_cited
  EditBox 100, 45, 160, 15, Claim_number
  EditBox 120, 65, 140, 15, months_of_overpayment
  EditBox 65, 85, 60, 15, discovery_date
  EditBox 200, 85, 60, 15, established_date
  EditBox 100, 105, 160, 15, reason_for_OP
  EditBox 150, 125, 110, 15, reason_to_be_reported
  EditBox 85, 145, 175, 15, supporting_docs
  EditBox 125, 165, 135, 15, responsible_parties
  EditBox 60, 185, 200, 15, total_amt_of_OP
  EditBox 70, 205, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 80, 240, 50, 15
    CancelButton 135, 240, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 115, 10, "Program(s) overpayment cited for:"
  Text 5, 70, 110, 10, "Month(s)/Year(s) of overpayment:"
  Text 5, 90, 55, 10, "Discovery date:"
  Text 135, 90, 60, 10, "Established date:"
  Text 5, 110, 95, 10, "Reason for OP (Be Specific):"
  Text 5, 150, 80, 10, "Supporting docs/verifs:"
  Text 5, 170, 120, 10, "Responsible parties listed by name:"
  Text 5, 190, 55, 10, "Total amt of OP:"
  Text 5, 210, 65, 10, "Sign the case note:"
  Text 130, 205, 125, 30, text_for_dialog
  Text 5, 50, 95, 10, "Claim Number(s) if available: "
  Text 5, 130, 140, 10, "When/why should this have been reported: "
EndDialog


'SECTION 03: THE SCRIPT

EMConnect ""


call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""



Do
  Do
    Do
      Dialog overpayment_dialog
      If buttonpressed = 0 then stopscript
      If case_number = "" then MsgBox "You must have a case number to continue!"
    Loop until case_number <> ""
    transmit
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You appear to be locked out of MAXIS. Are you passworded out? Did you navigate away from MAXIS?"
  Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
  call navigate_to_screen("case", "note")
  PF9
  EMReadScreen mode_check, 7, 20, 3
  If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "For some reason, the script can't get to a case note. Did you start the script in inquiry by mistake? Navigate to MAXIS production, or shut down the script and try again."
Loop until mode_check = "Mode: A" or mode_check = "Mode: E"

If Claim_number = "" Then
	Claim_number = "Not available at this time"
end if

call write_new_line_in_case_note("**OVERPAYMENT/CLAIM ESTABLISHED**")
call write_editbox_in_case_note("Program(s) overpayment cited for", programs_cited, 6) 
call write_editbox_in_case_note("Claim Number(s)", Claim_number, 6) 
call write_editbox_in_case_note("Month(s) of overpayment", months_of_overpayment, 6) 
call write_editbox_in_case_note("Discovery date", discovery_date, 6) 
call write_editbox_in_case_note("Established date", established_date, 6) 
call write_editbox_in_case_note("Reason for overpayment", reason_for_OP, 6) 
call write_editbox_in_case_note("When/Why should this have been reported", reason_to_be_reported, 6) 
call write_editbox_in_case_note("Supporting documents/verifications", supporting_docs, 6) 
call write_editbox_in_case_note("Responsible parties", responsible_parties, 6) 
call write_editbox_in_case_note("Total overpayment amount", total_amt_of_OP, 6) 
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

script_end_procedure("")
