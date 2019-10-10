
' Reusable Action  : Login 
' PURPOSE  : Login to the P+_menu applictaion
' Description:  The action invokes the application, and logs into the system, or resets the main window if the application is already up.
' Parameters : username, password
' Designby :   Azzi (Feb04,2019)
' Modified by:
' Comments

'==========================================================
 
SheetName="TestData"
'msgbox Parameter("sEndtoEndTestFilepath")
'sEndtoEndTestFilepath1 = Parameter("sEndtoEndTestFilepath")
DataTable.AddSheet(SheetName)
DataTable.ImportSheet Parameter("sEndtoEndTestFilepath"),SheetName,SheetName 
rowcount = DataTable.GetSheet(SheetName).GetRowCount
'msgbox rowcount		 ' Displays 3
DataTable.SetCurrentRow(1) 
UserName=Datatable.Value("UserName",SheetName)  
Password=Datatable.Value("Password",SheetName)  
'msgbox UserName
'msgbox Password
	
Browser("Login").Page("Login").WebEdit("username").Set UserName
Browser("Login").Page("Login").WebEdit("password").Set Password @@ script infofile_;_ZIP::ssf2.xml_;_
Browser("Login").Page("Login").WebEdit("username").Set "npp-enterprise" @@ script infofile_;_ZIP::ssf3.xml_;_
Browser("Login").Page("Login").WebElement("New Power+ Username Password").Click @@ script infofile_;_ZIP::ssf4.xml_;_
Browser("Login").Page("Login").WebButton("Sign-in").Click @@ script infofile_;_ZIP::ssf5.xml_;_

If true Then		
	  ' Reporter.ReportEvent micPass,1 ,"Login page appeared, credentials entered"
		LogResult micPass , "Login Page Should Appear" , "Login Page appeared Successfully"   
    Else
    	'Reporter.ReportEvent micFail, "Login Page Should Appear" ,"Login page did NOT Appeared"
    	LogResult micFail , "Login Page Should Appear" , "Login page did NOT Appeared"
    	testCleanUp()
 		ExitTest
	End If
	
'==========================================================
