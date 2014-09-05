'This is the starting point for the Renewals script. It is a hub from which additional scripts are run depending on--------------------------------- 
'which renewal paperwork is turned in.--------------------------------------------------------------------------------------------------------------

'Function(s)------------------------------------------------------------------------------------------------------
'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\County beta staging\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'This function allows this script to run other renewal-specific scripts ------------------------------------------
function run_renewals_script(renewals_script_path)
  Set run_another_renewals_script_fso = CreateObject("Scripting.FileSystemObject")
  Set fso_renewals_command = run_another_renewals_script_fso.OpenTextFile(renewals_script_path)
  text_from_the_other_renewals_script = fso_renewals_command.ReadAll
  fso_renewals_command.Close
  Execute text_from_the_other_renewals_script
  stopscript
end function

'Dialog(s)-------------------------------------------------------------------------------------------------------
BeginDialog renewal_dialog, 0, 0, 71, 115, "Renewal Dialog"
  ButtonGroup ButtonPressed
    PushButton 5, 5, 60, 10, "Combined AR", Combined_AR_button
    PushButton 5, 20, 60, 10, "CSR", CSR_button
    PushButton 5, 35, 60, 10, "HC ER", HC_ER_button
    PushButton 5, 50, 60, 10, "HRF (Family)", HRF_family_button
    CancelButton 5, 95, 60, 15
EndDialog


'The script--------------------------------------------------------------------------------------------------------------
Dialog renewal_dialog
IF buttonpressed = 0 THEN STOPSCRIPT

'Depending on the renewal paperwork that is submitted by the client, the script will open a new script----------------------------------
IF buttonpressed = CSR_button THEN
  run_renewals_script("Q:\Blue Zone Scripts\County beta staging\NOTE - CSR.vbs")
  STOPSCRIPT
END IF

IF buttonpressed = Combined_AR_button THEN
  run_renewals_script("Q:\Blue Zone Scripts\County beta staging\NOTE - Combined AR.vbs")
  STOPSCRIPT
END IF

IF buttonpressed = HC_ER_button THEN
  run_renewals_script("Q:\Blue Zone Scripts\County beta staging\NOTE - HC ER.vbs")
  STOPSCRIPT
END IF

IF buttonpressed = HRF_family_button THEN
  run_renewals_script("Q:\Blue Zone Scripts\County beta staging\NOTE - HRF (Family).vbs")
  STOPSCRIPT
END IF
