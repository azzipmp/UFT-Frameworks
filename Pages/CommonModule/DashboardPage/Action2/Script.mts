
' Reusable Action  : Select Customer
' PURPOSE  :   Click on the customer option in dashboard page
' Description:  The action click the custome option, post login
' Parameters : NA
' Designby :   Azzi (Feb04,2019)
' Modified by:
' Comments

'==========================================================
 
Browser("Dashboard").Page("Dashboard").WebElement("PFG").Click @@ script infofile_;_ZIP::ssf1.xml_;_
Browser("Dashboard").Page("Dashboard").WebButton("Okay").Click @@ script infofile_;_ZIP::ssf2.xml_;_
Browser("Dashboard").Page("Dashboard").WebElement("WebElement").Click @@ script infofile_;_ZIP::ssf3.xml_;_
Browser("Dashboard").Page("Dashboard").WebElement("WebElement_2").Click @@ script infofile_;_ZIP::ssf4.xml_;_
Browser("Dashboard").Page("Dashboard").WebButton("Okay").Click @@ script infofile_;_ZIP::ssf5.xml_;_
'Reporter.ReportEvent micPass,1 ,"Click Customer Option in Dashboard Page","Selected Customer Option in Dashboard Page"
LogResult micPass , "Click Customer Option in Dashboard Page" , "Selected Customer Option in Dashboard Page"
  
'========================================================================= @@ script infofile_;_ZIP::ssf7.xml_;_
