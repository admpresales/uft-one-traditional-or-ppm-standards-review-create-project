'===========================================================
'20201007 - DJ: Initial creation
'20201014 - DJ: Turned off smart identification for the final status object
'===========================================================

'===========================================================
'Function to search for the PPM proposal in the appropriate status
'===========================================================
Function PPMProposalSearch (CurrentStatus, NextAction)
	'===========================================================================================
	'BP:  Click the Search menu item
	'===========================================================================================
	Browser("Search Requests").Page("Dashboard - PFM Overview").Link("SEARCH").Click
	
	'===========================================================================================
	'BP:  Click the Requests text
	'===========================================================================================
	Browser("Search Requests").Page("Dashboard - PFM Overview").Link("Requests").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Enter PFM - Proposal into the Request Type field
	'===========================================================================================
	Browser("Search Requests").Page("Search Requests").WebEdit("Request Type Field").Set "PFM - Proposal"
	Browser("Search Requests").Page("Search Requests").WebElement("Status Label").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Enter a status of "New" into the Status field
	'===========================================================================================
	Browser("Search Requests").Page("Search Requests").WebEdit("Status Field").Set CurrentStatus
	
	'===========================================================================================
	'BP:  Click the Search button (OCR not seeing text, use traditional OR)
	'===========================================================================================
	Browser("Search Requests").Page("Search Requests").Link("Search").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Click the first record returned in the search results
	'===========================================================================================
	DataTable.Value("dtFirstReqID") = Browser("Search Requests").Page("Request Search Results").WebTable("Req #").GetCellData(2,2)
	Browser("Search Requests").Page("Request Search Results").Link("First Request ID Link").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	
End Function

Dim BrowserExecutable, Counter

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
BrowserExecutable = DataTable.Value("BrowserName") & ".exe"
SystemUtil.Run BrowserExecutable,"","","",3													'launch the browser specified in the data table
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon
Set AppContext2=Browser("CreationTime:=1")													'Set the variable for what application (in this case the browser) we are acting upon

'===========================================================================================
'BP:  Navigate to the PPM Launch Pages
'===========================================================================================

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("URL")													'Navigate to the application URL
AppContext.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Strategic Portfolio link
'===========================================================================================
Browser("Search Requests").Page("Project & Portfolio Management").Image("Strategic Portfolio Link").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Jonathan Kaplan (Portfolio Manager) link to log in as Jonathan Kaplan
'===========================================================================================
Browser("Search Requests").Page("Portfolio Management").WebArea("Jonathan Kaplan Image").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Search for proposals in a status of "Standards Review"
'===========================================================================================
PPMProposalSearch "Standards Review", "Status: Standards Review"

'===========================================================================================
'BP:  Click the left Approved button
'===========================================================================================
Browser("Search Requests").Page("Req Details").Link("First Approved Button").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the remaining Approved button
'===========================================================================================
Browser("Search Requests").Page("Req Details").Link("Approved Button").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the remaining Approved button
'===========================================================================================
Browser("Search Requests").Page("Req Details").Link("Approved Button").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Set the Project Manager to be Joseph Banks
'===========================================================================================
Browser("Search Requests").Page("Req More Information").WebEdit("Project Manager").Set "Joseph Banks"

'===========================================================================================
'BP:  Enter Standard Project (PPM) - Medium Size into the Projec Type field
'===========================================================================================
Browser("Search Requests").Page("Req More Information").WebEdit("Project Type Field").Set "Standard Project (PPM) - Medium Size"

'===========================================================================================
'BP:  Click the Continue Workflow Action button
'===========================================================================================
Browser("Search Requests").Page("Req More Information").WebElement("Continue Workflow Action Button").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Execute Now button
'===========================================================================================
Browser("Search Requests").Page("Req Details").Link("Execute Now").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Wait for the Status:Closed (Approved) to show up
'===========================================================================================
Counter = 0
Do
	Counter = Counter + 1
	wait(1)
	If Counter >=90 Then
		msgbox("Something is broken, status of the request hasn't shown up to be approved.")
		Reporter.ReportEvent micFail, "Create Project", "The project creation didn't finish within " & Counter & " exists timeouts."
		Exit Do
	End If
Loop Until Browser("Search Requests").Page("Req Details").WebElement("Status: Closed (Approved)").Exist(1)
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Logout
'===========================================================================================
Browser("Search Requests").Page("Req Details").WebElement("menuUserIcon").Click
AppContext.Sync																				'Wait for the browser to stop spinning
Browser("Search Requests").Page("Req Details").Link("Sign Out Link").Click
AppContext.Sync																				'Wait for the browser to stop spinning

AppContext.Close																			'Close the application at the end of your script

