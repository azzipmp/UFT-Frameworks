gblTakePassFailScreenshot = True '''If the Value is True then script will take Pass and Fail both steps Screenshot
CAPTURE_ERROR_SCREENSHOT = True	  '''And IF value is False then Script will take only Fail steps screenshots
Public slNo
' ***********************************************************************************************
'
' 			R E P O R T   L I B R A R Y   F U N C T I O N S 
'
' ***********************************************************************************************
'
'1. CreateResultFile
'2. LogResult
'3. TestCaseExecutiveSummary

' ================================================================================================
'  NAME			    : CreateResultFile
'  DESCRIPTION 	  	: This function creates HTML test script execution result. Result file name 
'			           created is same as the test script name. This function is called at the 
'			           begining of the test script
' ================================================================================================

Set objBrowser = Description.Create
Set objWindow  = Description.Create
objBrowser("creationtime").value = "0"
objWindow("regexpwndtitle").value = "Microsoft Internet Explorer|Windows Internet Explorer|Mozilla Firefox"

Public iPassCount, iFailCount, sResultFile, iStartTime, iErrImageNumber

iPassCount = 0 
iFailCount = 0
iErrImageNumber = 1

Public Function CreateResultFile(sTestName)
	Dim fso, MyFile, i
	iStartTime = Now
	slNo = 0              'initialise serial number of testcase
	'RelPath = Pathfinder.locate("..\..\")
	 RelPath=Pathfinder.locate("..\")
	' msgbox RelPath
' sResultFile = Relpath & "\TestResults\"& Environment.Value("TestName") & "_" & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & "_" & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now) & ".html"
 sResultFile = Relpath & "TestResults\"& sTestName & "_" & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & "_" & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now) & ".html"
'msgbox sResultFile
	Set fso = CreateObject("Scripting.FileSystemObject")
	' Create Module folder if it doesn't exists
'	msgbox strResultPath
'	If (fso.FolderExists (strResultPath) = False) Then
'		Set fFolder = fso.CreateFolder (strResultPath)
'	End If
		
	' Create Test Result file from TestName 
'	sResultFile = strResultPath & Environment.Value("TestName") & "_" & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & "_" & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now) & ".html"
'	msgbox sResultFile
	Set MyFile = fso.CreateTextFile(sResultFile,True)
	
	MyFile.WriteLine("<html><head><title>Test Script Execution Report</title></head>")
	MyFile.WriteLine("<body><table border='1' width='100%' bordercolorlight='#C0C0C0' cellspacing='0' cellpadding='0' bordercolordark='#C0C0C0' bordercolor='#C0C0C0' height='88'>")
	MyFile.WriteLine("<tr><td width='100%' colspan='4' height='28' bgcolor='#C0C0C0'><p align='center'><b><font face='Verdana' size='4' color='#000080'>Test Script Execution Report</font></b></td></tr>")
	MyFile.WriteLine("<tr><td width='16%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Test Script Path</font></b></td>")
	MyFile.WriteLine("<td width='84%' height='25' colspan='3'><p style='margin-left: 5'><font face='Verdana' size='2'>")
	
	MyFile.WriteLine(Environment.Value("TestDir") & "</font></td></tr>")
	MyFile.WriteLine("<tr><td width='16%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Test Case Name</font></b></td>")
	MyFile.WriteLine("<td width='84%' height='25' colspan='3'><font face='Verdana' size='2'>&nbsp;" & Environment.Value("TestName") & "</font></td></tr>	")
	
	MyFile.WriteLine("<tr><td width='16%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Test URL</font></b></td>")
	MyFile.WriteLine("<td width='84%' height='25' colspan='3'><font face='Verdana' size='2'>&nbsp;"& strURL &"</font></td></tr>	")
	
	MyFile.WriteLine("</table>")
	MyFile.WriteLine("<p style='margin-left: 5'>&nbsp; </p>")
	
	MyFile.WriteLine("<table border='1' width='100%' bordercolordark='#C0C0C0' cellspacing='0' cellpadding='0' bordercolorlight='#C0C0C0' bordercolor='#C0C0C0' height='91'>")	
	
	'============== Result Column Header =====================
	MyFile.WriteLine("<tr><td width='5%' align='center' bgcolor='#000099' height='35'><b>")
	MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Sl. No.</font></b></td>")
	MyFile.WriteLine("<td width='45%' align='center' bgcolor='#000099' height='35'><b>")
	''MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Step Description</font></b></td>")
	MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Test Step Description/Expected Result</font></b></td>")
	MyFile.WriteLine("<td width='40%' align='center' bgcolor='#000099' height='35'><b>")
	''MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Details</font></b></td>")
	MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Test Step Actual Result</font></b></td>")
	MyFile.WriteLine("<td width='10%' bgcolor='#000099' height='35' align='center'><b>")
	''MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Status</font></b></td></tr>")
	MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Test Step Status</font></b></td></tr>")
			
	' Close the file
	MyFile.Close
	CreateResultFile = resultfile
End Function

' ================================================================================================
'  NAME				: LogResult
'  DESCRIPTION 	  	: This function is used to write test step status to QTP result file and 
'					  also to the HTML test result file.
'  PARAMETERS		: sStatus 	- Status of the step (micPass / micFail / micGeneral)
'				  sTestStep 	- Test Step Name
'	       		  sDescription 	- Test Step Description
' ================================================================================================

Public Function LogResult(sStatus, sTestStep, sDescription)
	Const ForAppending = 8
	Const TristateUseDefault = -2
	Dim fso, f, ts
	Set fso = CreateObject("Scripting.FileSystemObject")
	If (fso.FolderExists (strResultPath & "ErrorSnapshot") = False) Then
		fso.CreateFolder (strResultPath & "ErrorSnapshot")
	End If
	
	Set f = fso.GetFile(sResultFile)
	Set ts = f.OpenAsTextStream(ForAppending, TristateUseDefault)
	
	If (sStatus = micGeneral) Then
		Reporter.ReportEvent micGeneral, sTestStep, sDescription
		ts.WriteLine("<tr><td width='12%' height='25' align='left' colspan='3'><p style='margin-left: 5'><b><font face='Verdana' size='2' color='#000099'> " & sTestStep & " " & sDescription & "</font></b></td></tr>")
        ts.Close
        Exit Function
	End If
	slNo = slNo + 1
	ts.WriteLine("<tr><td width='5%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & slNo & "</font></td>")
	ts.WriteLine("<td width='45%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & sTestStep & "</font></td>")
	ts.WriteLine("<td width='40%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>")
	
	' Replace vbcrlf in sDescription to <br>
	arrDesc = Split (sDescription, vbcrlf)
	
	For Each arrElement In arrDesc
		ts.WriteLine(arrElement & "<br>")
	Next
	ts.WriteLine("</font></td>")
    If sStatus = micPass Then
    	Reporter.ReportEvent micPass, sTestStep, sDescription
    	''If (gblTakePassFailScreenshot = True ) and (iPassCount <> 0 or iPassCount <> 1) Then
    	If (gblTakePassFailScreenshot = True ) Then
    	RelPath = Pathfinder.locate("..\")
 
			sImgRelativePath = "ErrorSnapshot\" & Environment.Value("TestName") & "_" & Month(Now) & "_" & Day(Now) & "_" & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now) & "_PASS_STEP_" & iErrImageNumber & ".png"
			sErrImage1 = Relpath & "TestResults\"&sImgRelativePath
			Desktop.CaptureBitmap sErrImage1, True		' Capture Desktop Snapshot.So that the browser will be displayed with its Address bar, which we miss by Browser().CaptureBitmap If image by name sErrImage already exist, then override		
			ts.WriteLine("<td width='10%' height='25' align='center'><p style='margin-left: 5'><font face='Verdana' size='2' color='#05A251'><a href='" & sImgRelativePath & "'>PASS </a></font></td></tr>")	    		
		Else
			ts.WriteLine("<td width='10%' height='25' align='center'><p style='margin-left: 5'><b><font face='Verdana' size='2' color='#05A251'> PASS </font></b></td></tr>")
		End If
        ts.Close
        iPassCount = iPassCount + 1
	ElseIf sStatus = micFail Then
		Reporter.ReportEvent micFail, sTestStep, sDescription
		If (Browser(objBrowser).Exist(5)) Then
			If (CAPTURE_ERROR_SCREENSHOT = True) Then 
				sImgRelativePath = "ErrorSnapshot\" & Environment.Value("TestName") & "_" & Month(Now) & "_" & Day(Now) & "_" & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now) & "_FAIL_STEP_" & iErrImageNumber & ".png"
				sErrImage1 = strResultPath & sImgRelativePath
				Desktop.CaptureBitmap sErrImage1, True		' Capture Desktop Snapshot.So that the browser will be displayed with its Address bar, which we miss by Browser().CaptureBitmap If image by name sErrImage already exist, then override
				ts.WriteLine("<td width='10' height='25' align='center'><p style='margin-left: 5'><font face='Verdana' size='2' color='#FF0000'><a href='" & sImgRelativePath & "'>FAIL </a></font></td></tr>")
				iErrImageNumber = iErrImageNumber + 1
			Else
				ts.WriteLine("<td width='10%' height='25' align='center'><p style='margin-left: 5'><font face='Verdana' size='2' color='#FF0000'>FAIL</font></td></tr>")
			End If
		Else
			ts.WriteLine("<td width='10%' height='25' align='center'><p style='margin-left: 5'><font face='Verdana' size='2' color='#FF0000'>FAIL</font></td></tr>")
		End If
		iFailCount = iFailCount + 1
	    ts.Close
	End If
End Function

' ================================================================================================
'  NAME			: TestCaseExecutiveSummary
'  DESCRIPTION 	  	: This function is used to create test script executive summary. This function 
'					  is called from test script at the end of the test script
'  PARAMETERS		:
' ================================================================================================

Function TestCaseExecutiveSummary ()
	Const ForAppending = 8
	Const TristateUseDefault = -2
	Dim fso, f, ts
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFile(sResultFile)
	Set ts = f.OpenAsTextStream(ForAppending, TristateUseDefault)

	ts.WriteLine("</table>")
	ts.WriteLine("<p>&nbsp;</p>")
	ts.WriteLine("<table border='1' width='52%' bordercolorlight='#C0C0C0' cellspacing='0' cellpadding='0' bordercolordark='#C0C0C0' bordercolor='#C0C0C0' height='88'>")
	ts.WriteLine("<tr><td width='53%' colspan='2' height='28' bgcolor='#C0C0C0'><p align='center'><b><font face='Verdana' size='4' color='#000080'>")
	ts.WriteLine("Test Script Execution Summary</font></b></td></tr>")
	
	ts.WriteLine("<tr><td width='17%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2' color='#05A251'>")
	ts.WriteLine("Total Pass Count</font></b></td>")
	
	ts.WriteLine("<td width='36%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'color='#05A251'>&nbsp;" & iPassCount & "</font></td></tr>")
	ts.WriteLine("<tr> <td width='17%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2' color='#FF0000'>")
	ts.WriteLine("Total Fail Count</font></b></td>")
	
	ts.WriteLine("<td width='36%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'color='#FF0000'>&nbsp;" & iFailCount & "</font></td></tr>")
	ts.WriteLine("<tr><td width='17%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>")
	ts.WriteLine("Start Time</font></b></td>")
	ts.WriteLine("<td width='36%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & iStartTime & "</font></td></tr>")
	ts.WriteLine("<tr> <td width='17%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>End Time</font></b></td>")
	ts.WriteLine("<td width='36%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & Now & "</font></td></tr></table>")
	ts.Close
	
	' If iFailCount = 0, then set TestRunResults status to pass by executing Reporter.ReportEvent micPass statement
	Dim fso_TestScriptStatus
	Set fso_TestScriptStatus = CreateObject("Scripting.FileSystemObject")
	Set TestScriptExecutionStatus = fso_TestScriptStatus.CreateTextFile(strLibPath & "\TCStatus.txt", True)
	If (iFailCount = 0) Then
		TestScriptExecutionStatus.WriteLine ("Passed")
		'sTCRunStatus = "Passed"
	Else
		TestScriptExecutionStatus.WriteLine ("Failed")
		'sTCRunStatus = "Failed"
	End If
	TestScriptExecutionStatus.Close
	
'	testDir = createobject("Scripting.Filesystemobject").GetAbsolutePathName(".")
'	testSuitControllerFile = testDir & "\" & "Controller.xlsx"
'	Set objContExcel = CreateObject("Excel.Application")
'	objContExcel.Workbooks.Open(testSuitControllerFile)
'	Set objWorkbook = objContExcel.ActiveWorkbook.Worksheets("Controller")
'	If (sTCRunStatus = "Failed") Then
'		objWorkbook.Cells(intRow, 5).Value = ""
'		objWorkbook.Cells(intRow, 5).Value = "FAIL"
'	ElseIf (sTCRunStatus = "Passed") Then
'		objWorkbook.Cells(intRow, 5).Value = ""			
'		objWorkbook.Cells(intRow, 5).Value = "PASS"
'	Else
'		objWorkbook.Cells(intRow, 5).Value = ""		
'		objWorkbook.Cells(intRow, 5).Value = sTCRunStatus
'	End If
'			
	Set TestScriptExecutionStatus 	= Nothing
	Set fso_TestScriptStatus 		= Nothing
	systemutil.CloseProcessByName("EXCEL.EXE")
End Function 
