'----------This is a test of a request for a SIR Enhancement Request to pull 
'----------The intent of this script is to read STAT panels and print them to Word, 2 to a page.
'----------The script can check up to 39 panels, although that number can be adjusted by adjusting the dialog.

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'----------FUNCTIONS----------
FUNCTION stat_panel_to_word(x)
	EMWriteScreen left(x, 4), 20, 71
	EMWriteScreen left(right(x, 5), 2), 20, 76
	EMWriteScreen right(x, 2), 20, 79
	transmit

	EMReadScreen panel_does_not_exist, 14, 24, 13
		IF panel_does_not_exist <> "DOES NOT EXIST" THEN
			stat_row = 1
			DO
				EMReadScreen stat_line, 80, stat_row, 1
				stat_line = replace(stat_line, " ", "_")
				page_array = page_array & stat_line & " "
				stat_row = stat_row + 1
			LOOP UNTIL stat_row = 25
		ELSE
			stat_line = x & " DOES NOT EXIST"
			page_array = page_array & stat_line & " "
		END IF
		
		page_array = trim(page_array)
		page_array = split(page_array)

		FOR EACH panel_read IN page_array
			panel_read = replace(panel_read, "_", " ")
			objselection.typetext(panel_read)
			objselection.TypeParagraph()
		NEXT
END FUNCTION

'----------DIALOGS----------
BeginDialog read_panel_dialog, 0, 0, 241, 345, "Select STAT Panels"
  EditBox 75, 5, 60, 15, case_number
  EditBox 10, 55, 70, 15, panel01
  EditBox 10, 75, 70, 15, panel02
  EditBox 10, 95, 70, 15, panel03
  EditBox 10, 115, 70, 15, panel04
  EditBox 10, 135, 70, 15, panel05
  EditBox 10, 155, 70, 15, panel06
  EditBox 10, 175, 70, 15, panel07
  EditBox 10, 195, 70, 15, panel08
  EditBox 10, 215, 70, 15, panel09
  EditBox 10, 235, 70, 15, panel10
  EditBox 10, 255, 70, 15, panel11
  EditBox 10, 275, 70, 15, panel12
  EditBox 10, 295, 70, 15, panel13
  EditBox 85, 55, 70, 15, panel14
  EditBox 85, 75, 70, 15, panel15
  EditBox 85, 95, 70, 15, panel16
  EditBox 85, 115, 70, 15, panel17
  EditBox 85, 135, 70, 15, panel18
  EditBox 85, 155, 70, 15, panel19
  EditBox 85, 175, 70, 15, panel20
  EditBox 85, 195, 70, 15, panel21
  EditBox 85, 215, 70, 15, panel22
  EditBox 85, 235, 70, 15, panel23
  EditBox 85, 255, 70, 15, panel24
  EditBox 85, 275, 70, 15, panel25
  EditBox 85, 295, 70, 15, panel26
  EditBox 160, 55, 70, 15, panel27
  EditBox 160, 75, 70, 15, panel28
  EditBox 160, 95, 70, 15, panel29
  EditBox 160, 115, 70, 15, panel30
  EditBox 160, 135, 70, 15, panel31
  EditBox 160, 155, 70, 15, panel32
  EditBox 160, 175, 70, 15, panel33
  EditBox 160, 195, 70, 15, panel34
  EditBox 160, 215, 70, 15, panel35
  EditBox 160, 235, 70, 15, panel36
  EditBox 160, 255, 70, 15, panel37
  EditBox 160, 275, 70, 15, panel38
  EditBox 160, 295, 70, 15, panel39
  ButtonGroup ButtonPressed
    OkButton 70, 325, 50, 15
    CancelButton 125, 325, 50, 15
  Text 10, 10, 55, 10, "Case Number"
  Text 10, 30, 225, 20, "Panels to select (please format as you would on the Command line - for example JOBS/01/01)"
EndDialog

'----------THE SCRIPT----------
EMConnect ""

maxis_check_function

DIALOG read_panel_dialog
	IF ButtonPressed = 0 THEN stopscript

	IF panel01 <> "" THEN panel_array = panel_array & panel01 & " "
	IF panel02 <> "" THEN panel_array = panel_array & panel02 & " "
	IF panel03 <> "" THEN panel_array = panel_array & panel03 & " "
	IF panel04 <> "" THEN panel_array = panel_array & panel04 & " "
	IF panel05 <> "" THEN panel_array = panel_array & panel05 & " "
	IF panel06 <> "" THEN panel_array = panel_array & panel06 & " "
	IF panel07 <> "" THEN panel_array = panel_array & panel07 & " "
	IF panel08 <> "" THEN panel_array = panel_array & panel08 & " "
	IF panel09 <> "" THEN panel_array = panel_array & panel09 & " "
	IF panel10 <> "" THEN panel_array = panel_array & panel10 & " "
	IF panel11 <> "" THEN panel_array = panel_array & panel11 & " "
	IF panel12 <> "" THEN panel_array = panel_array & panel12 & " "
	IF panel13 <> "" THEN panel_array = panel_array & panel13 & " "
	IF panel14 <> "" THEN panel_array = panel_array & panel14 & " "
	IF panel15 <> "" THEN panel_array = panel_array & panel15 & " "
	IF panel16 <> "" THEN panel_array = panel_array & panel16 & " "
	IF panel17 <> "" THEN panel_array = panel_array & panel17 & " "
	IF panel18 <> "" THEN panel_array = panel_array & panel18 & " "
	IF panel19 <> "" THEN panel_array = panel_array & panel19 & " "
	IF panel20 <> "" THEN panel_array = panel_array & panel20 & " "
	IF panel21 <> "" THEN panel_array = panel_array & panel21 & " "
	IF panel22 <> "" THEN panel_array = panel_array & panel22 & " "
	IF panel23 <> "" THEN panel_array = panel_array & panel23 & " "
	IF panel24 <> "" THEN panel_array = panel_array & panel24 & " "
	IF panel25 <> "" THEN panel_array = panel_array & panel25 & " "
	IF panel26 <> "" THEN panel_array = panel_array & panel26 & " "
	IF panel27 <> "" THEN panel_array = panel_array & panel27 & " "
	IF panel28 <> "" THEN panel_array = panel_array & panel28 & " "
	IF panel29 <> "" THEN panel_array = panel_array & panel29 & " "
	IF panel30 <> "" THEN panel_array = panel_array & panel30 & " "
	IF panel31 <> "" THEN panel_array = panel_array & panel31 & " "
	IF panel32 <> "" THEN panel_array = panel_array & panel32 & " "
	IF panel33 <> "" THEN panel_array = panel_array & panel33 & " "
	IF panel34 <> "" THEN panel_array = panel_array & panel34 & " "
	IF panel35 <> "" THEN panel_array = panel_array & panel35 & " "
	IF panel36 <> "" THEN panel_array = panel_array & panel36 & " "
	IF panel37 <> "" THEN panel_array = panel_array & panel37 & " "
	IF panel38 <> "" THEN panel_array = panel_array & panel38 & " "
	IF panel39 <> "" THEN panel_array = panel_array & panel39 & " "

panel_array = trim(panel_array)
panel_array = split(panel_array)

Set objWord = CreateObject("Word.Application")
objWord.Visible = true
set objDoc = objWord.Documents.add()
Set objSelection = objWord.Selection
Set objRange = objDoc.Range()
objRange.Font.Name = "Lucida Console"
objRange.Font.Size = 8

FOR EACH stat_panel IN panel_array
	call stat_panel_to_word(stat_panel)
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
NEXT

