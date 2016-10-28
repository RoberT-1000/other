
BeginDialog Dialog1, 0, 0, 236, 325, "Dialog"
  EditBox 105, 50, 120, 15, integrated_case_number
  EditBox 105, 70, 120, 15, recipient_id
  EditBox 105, 90, 120, 15, mnsure_id
  EditBox 105, 110, 120, 15, first_name
  EditBox 105, 130, 120, 15, last_name
  EditBox 105, 150, 120, 15, date_of_birth
  EditBox 105, 170, 120, 15, soc_sec_num
  EditBox 105, 190, 120, 15, smi_number
  EditBox 105, 210, 120, 15, addrOne
  EditBox 105, 230, 120, 15, addrTwo
  EditBox 105, 250, 50, 15, city
  EditBox 160, 250, 20, 15, state
  EditBox 185, 250, 40, 15, zip
  ButtonGroup ButtonPressed
    OkButton 125, 300, 50, 15
    CancelButton 175, 300, 50, 15
    PushButton 5, 300, 60, 15, "About", about_button
  Text 10, 10, 215, 25, "Please complete the following fields. Anything with an ASTERISK is required. Additional information can be found by pressing the ''About'' button."
  Text 10, 55, 90, 10, "Integrated Case Number*"
  Text 10, 75, 85, 10, "Recipient ID/PMI*"
  Text 10, 95, 85, 10, "MNsure ID*"
  Text 10, 115, 85, 10, "First Name*"
  Text 10, 135, 85, 10, "Last Name*"
  Text 10, 155, 85, 10, "Date of Birth*"
  Text 10, 175, 85, 10, "Social Security Number*"
  Text 10, 195, 85, 10, "SMI Number"
  Text 10, 215, 85, 10, "Address*"
EndDialog

DO
	err_msg = ""
	Dialog
		IF ButtonPressed = 0 THEN stopscript
		IF ButtonPressed = about_button THEN 
			SET aboutMsg = CreateObject("Scripting.FileSystemObject")
			SET aboutMsgCmd = aboutMsg.OpenTextFile("Q:\Blue Zone Scripts\MNsure Duplicate People\about me.txt")
			aboutMsgTxt = aboutMsgCmd.ReadAll
			aboutMsgCmd.Close

			MsgBox aboutMsgTxt
		END IF
		'validating the values in the edit boxes
		'...integrated case number
		IF ButtonPressed = -1 THEN 
			IF integrated_case_number = "" THEN 
				err_msg = err_msg & vbCr & "* Please enter an Integrated Case Number."
			ELSEIF integrated_case_number <> "" AND IsNumeric(integrated_case_number) = FALSE THEN 
				err_msg = err_msg & vbCr & "* The Integrated Case Number you provided is not numeric. Please provide a valid, numeric Integrated Case Number."
			ELSEIF integrated_case_number <> "" AND IsNumeric(integrated_case_number) = TRUE AND len(integrated_case_number) <> 8 THEN 
				err_msg = err_msg & vbCr & "* The Integrated Case Number you provided is not eight digits long. Please provide a valid Integrated Case Number."
			END IF
			'...recipient ID
			IF recipient_id = "" THEN 
				err_msg = err_msg & vbCr & "* Please provide a Recipient ID/PMI."
			ELSEIF recipient_id <> "" AND IsNumeric(recipient_id) = FALSE THEN 
				err_msg = err_msg & vbCr & "* Please provide a valid, numeric Recipient ID/PMI."
			ELSEIF recipient_id <> "" AND IsNumeric(recipient_id) = TRUE AND len(recipient_id) <> 8 THEN 
				err_msg = err_msg & vbCr & "* Please provide a valid, 8-digit Recipient ID/PMI."
			END IF
			'...mnsure ID
			IF mnsure_id = "" THEN 
				err_msg = err_msg & vbCr & "* Please provide a MNsure ID."
			ELSEIF mnsure_id <> "" AND IsNumeric(mnsure_id) = FALSE THEN 
				err_msg = err_msg & vbCr & "* Please provide a valid, numeric MNsure ID."
			ELSEIF mnsure_id <> "" AND IsNumeric(mnsure_id) = TRUE AND len(mnsure_id) <> 10 THEN 
				err_msg = err_msg & vbCr & "* Please provide a valid, ten-digit MNsure ID."
			END IF
			IF first_name = "" THEN err_msg = err_msg & vbCr & "* Please provide a first name for this client."
			IF last_name = "" THEN err_msg = err_msg & vbCr & "* Please provide a last name for this client."
			'...date of birth
			IF date_of_birth = "" THEN 
				err_msg = err_msg & vbCr & "* Please provide a date of birth for this client."
			ELSEIF date_of_birth <> "" AND IsDate(date_of_birth) = FALSE THEN 
				err_msg = err_msg & vbCr & "* Please provide a valid date of birth (in a date format MM/DD/YYYY) for this client."
			END IF
			'...social
			IF soc_sec_num = "" THEN 
				err_msg = err_msg & vbCr & "* You must provide a Social Security Number for this client."
			ELSEIF len(soc_sec_num) <> 11 THEN 
				err_msg = err_msg & vbCr & "* Please provide a valid Social Security Number for this client in the XXX-XX-XXXX format."
			'trying this again on the SSN but this time replacing the dashes with blanks and seeing if the user entered 9 digits
			ELSEIF len(replace(soc_sec_num, "-", "")) <> 9 AND IsNumeric(replace(soc_sec_num, "-", "")) = FALSE THEN 
				err_msg = err_msg & vbCr & "* Please provide a valid Social Security Number for this client in the XXX-XX-XXXX format."
			END IF
			IF addrOne = "" THEN err_msg = err_msg & vbCr & "* Please provide a valid address for this client."
			IF city = "" THEN err_msg = err_msg & vbCr & "* Please provide the city of the client's address."
			IF state = "" THEN err_msg = err_msg & vbCr & "* Please provide the state of the client's address."
			IF zip = "" THEN 
				err_msg = err_msg & vbCr & "* Please provide a five-digit zip code for the client's address."
			ELSEIF zip <> "" AND len(zip) <> 5 THEN 
				err_msg = err_msg & vbCr & "* Please provide a five-digit zip code for the client's address."
			ELSEIF zip <> "" AND IsNumeric(zip) = FALSE THEN 
				err_msg = err_msg & vbCr & "* Please provide a five-digit zip code for the client's address."
			END IF
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		END IF
LOOP UNTIL err_msg = "" AND ButtonPressed = -1

'Setting constants
Const adOpenStatic = 3
Const adLockOptimistic = 3

duplicate_db_lan = "**************"

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

'Opening DB
objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "" & duplicate_db_lan & ""

'Opening usage_log and adding a record
objRecordSet.Open "INSERT INTO ******** (ServiceLocation, CountyofResidence, IntegratedCaseNumber, RecipientID, MNsureID, FirstName, MiddleInitial, LastName, MajorProgram, BeginDate, EndDate, MAXISCaseNumber, Program, Status, BirthDate, Gender, SSN, SMINumber, Address1, Address2, City, State, ZipCode)" & _
	"VALUES ('002', '002', '" & integrated_case_number & "', '" & recipient_id & "', '" & mnsure_id & "', '" & first_name & "', 'X', '" & last_name & "', 'MA', '" & date & "', '" & date & "', 'N/A', 'N/A', 'N/A', '" & date_of_birth & "', 'N/A', '" & soc_sec_num & "', '" & smi_number & "', '" & addrOne & "', '" & addrTwo & "', '" & city & "', '" & state & "', '" & zip & "')", objConnection, adOpenStatic, adLockOptimistic

'sending data to the ODBC connected table
Set objRecord = CreateObject("ADODB.Recordset")

'On Error Resume Next		'to void any error message
'this is the ODBC connected table
objRecord.Open "INSERT INTO ******** (MNSureID, MNSureCaseID, PMINumber, SMINumber, ClientSSN, ClientDOB, LastName, FirstName, MiddleName, Address1, Address2, City, State, ZipCode)" & _
	"VALUES ('" & mnsure_id & "', '" & integrated_case_number & "', '" & recipient_id & "', '" & smi_number & "', '" & soc_sec_num & "', '" & date_of_birth & "', '" & last_name & "', '" & first_name & "', '', '" & addrOne & "', '" & addrTwo & "', '" & city & "', '" & state & "', '" & zip & "')", objConnection, adOpenStatic, adLockOptimistic

SET objNet = CreateObject("wscript.network")
windows_user_id = UCASE(objNet.UserName)
	
'sending data to user table for to check on people
SET objUser = CreateObject("ADODB.RecordSet")
objUser.Open "INSERT INTO ******** (MNSureID, userID, sdate)" & _ 
	"VALUES ('" & mnsure_id & "', '" & windows_user_id & "', '" & date & "')", objConnection, adOpenStatic, adLockOptimistic

MsgBox "Finished"






