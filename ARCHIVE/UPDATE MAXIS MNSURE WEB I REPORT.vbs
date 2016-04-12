'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"

SET req = CreateObject("Msxml2.XMLHttp.6.0") 'Creates an object to get a URL
req.open "GET", url, FALSE	'Attempts to open the URL
req.send 'Sends request

IF req.Status = 200 THEN	'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject") 'Creates an FSO
	Execute req.responseText 'Executes the script code
ELSE	'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox "Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
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

EMConnect ""

start_time = timer

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("H:\Tech Analyst\MNSure MAXIS Clients 2015-09-01.xls")
objExcel.DisplayAlerts = True
objExcel.Worksheets("MX Case with Worker (dup rem'd)").Activate

On Error Resume Next
		
Set objExplorer = CreateObject("InternetExplorer.Application")
objExplorer.Navigate "about:blank"   
objExplorer.ToolBar = 0
objExplorer.StatusBar = 0
objExplorer.Width = 600
objExplorer.Height = 100 
objExplorer.Visible = 1             
objExplorer.Document.Title = "Cleaning up report."
objExplorer.Document.Body.InnerHTML = "The script is finding the MAXIS workers."
Wscript.Sleep 1

FOR i = 2 to 2692
	back_to_SELF
	
	maxis_case_number = objExcel.Cells(i, 14).Value
	x_number = ""
	supervisor = ""
	
	EMWriteScreen "CASE", 16, 43
	EMWriteScreen maxis_case_number, 18, 43
	transmit
	
	EMReadScreen at_self, 4, 2, 50
	IF at_self = "SELF" THEN 
		CALL find_variable("PRIVILEGED WORKER: ", x_number, 7)
		supervisor = "PRIVILEGED"
	ELSE
		EMReadScreen x_number, 7, 21, 16
		EMSetCursor 21, 16
		PF1
		EMReadScreen supervisor, 20, 22, 16
		supervisor = trim(supervisor)
		transmit
	END IF
	
	objExcel.Cells(i, 15).Value = x_number
	objExcel.Cells(i, 16).Value = supervisor
	
	current_time = timer
	run_time = current_time - start_time
	objExplorer.Document.Body.InnerHTML = "The script is finding the MAXIS workers. It is " & FormatPercent((i - 1)/2692) & " complete. Current run time = " & run_time & " seconds."
NEXT
MsgBox "Success!!"

