Public strResultPath, strScriptPath, strLibPath, strAppLibPath, strReportLibPath, strTestDataPath, strObjRepPath, strCommonLibPath, strContPath, strBrowser , strURL

on error resume next
' ***********************************************************************************************
'
' 			C O M M O N   L I B R A R Y   F U N C T I O N S 
'
' ***********************************************************************************************

'1.strFormatNowDate
'2.Initialization
'3.booleanDBConnection
'4.booleanDBQuery
'5.booleanDBDisconnect



' ================================================================================================
'  NAME			: strFormatNowDate
'  DESCRIPTION 	  	: This function returns a string with date-time stamp
'  PARAMETERS		: nil
' ================================================================================================
Function strFormatNowDate()
	d=now()
	Dim arr(6)
	arr(0)="TestResults"
	arr(1)=DatePart("d",d)
	arr(2)=DatePart("m",d)
	arr(3)=DatePart("yyyy",d)
	arr(4)=DatePart("h",d) & "h"
	arr(5)=DatePart("n",d) & "m"
	arr(6)=DatePart("s",d)	& "s"
	strFormatNowDate= Join(arr,"_")
End Function


' ================================================================================================
'  NAME			: Initialization
'  DESCRIPTION 	  	: This function is used to create global variables which stores location 
'			   path of TestResult, TestData, Scripts, AppLib, Browser CommonLib & ObjectRepo 
'			  Loads common repository
'  PARAMETERS		: nil
' ================================================================================================

Public Function Initialization ()
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}" & "!\\.\root\cimv2")
	Set colProcess = objWMIService.ExecQuery ("Select * From Win32_Process")
	For Each objProcess in colProcess
		If LCase(objProcess.Name) = LCase("EXCEL.EXE") OR LCase(objProcess.Name) = LCase("EXCEL.EXE *32") Then
        objProcess.Terminate()
        ' MsgBox "- ACTION: " & objProcess.Name & " terminated"
		End If
	Next
	For Each objProcess in colProcess
		If LCase(objProcess.Name) = LCase("WerFault.exe") OR LCase(objProcess.Name) = LCase("WerFault.exe *32") Then
        objProcess.Terminate()
        ' MsgBox "- ACTION: " & objProcess.Name & " terminated"
		End If
	Next
	
	
	
	'Get the Directory of the framework
	SystemUtil.CloseProcessByName "EXCEL.EXE"
	sTestDir= Environment.Value ("TestDir")
	arrPath = Split (sTestDir, "\")

	'Save Result Path to variable strResultPath
	arrPath(UBound(arrPath)-1) = "TestResults"
	For I=0 to UBound(arrPath)-1
		If (I=0) Then
			strResultPath = arrPath(I)
		Else
			strResultPath = strResultPath + "\" + arrPath(I)
		End If
	Next
	strResultPath = strResultPath & "\"
msgbox strResultPath

	'Save Script Path to variable strScriptPath
	arrPath(UBound(arrPath)-1) = "Scripts"
	For I=0 to UBound(arrPath) -1
		If (I=0) Then
			strScriptPath = arrPath(I)
		Else
			strScriptPath = strScriptPath + "\" + arrPath(I)
		End If
	Next
	strScriptPath = strScriptPath  & "\"
	
    	'Save Lib Path to variable sAppLibPath
	arrPath(UBound(arrPath)-1) = "Library"
	For I=0 to UBound(arrPath) -1
		If (I=0) Then
			strLibPath = arrPath(I)
		Else
			strLibPath = strLibPath + "\" + arrPath(I)
		End If
	Next
	
	'Save TestData Path to variable strTestDataPath
	arrPath(UBound(arrPath)-1) = "TestData"
	For I=0 to UBound(arrPath) -1
		If (I=0) Then
			strTestDataPath = arrPath(I)
		Else
			strTestDataPath = strTestDataPath + "\" + arrPath(I)
		End If
	Next
	strTestDataPath = strTestDataPath & "\"	
	
	'Save ObjectRepository Path to variable strObjRepPath
	arrPath(UBound(arrPath)-1) = "COR"
	For I=0 to UBound(arrPath)-1
		If (I=0) Then
			strObjRepPath = arrPath(I)
		Else
			strObjRepPath = strObjRepPath + "\" + arrPath(I)
		End If
	Next
	strObjRepPath1 = strObjRepPath & "\" & "CommonObjectRepository.tsr"
	
	'Loading the repository file
	If  Instr(1, UCase(Environment.Value ("TestName")), "WS") = 0 Then
		RepositoriesCollection.RemoveAll() 
		RepositoriesCollection.Add(strObjRepPath1)  
	End If
	
	'Loading the Controller File PathDriver Path
	arrPath(UBound(arrPath)-1) = "Controller.xlsx"
	For I=0 to UBound(arrPath)-1
		If (I=0) Then
			strContPath = arrPath(I)
		Else
			strContPath = strContPath + "\" + arrPath(I)
		End If
	Next
	
	'Get the Browser from Controller File matching the test name
	Set fsoTestSuiteController = CreateObject("Scripting.FileSystemObject")
	Set sTestSuiteControllerFileSpec = fsoTestSuiteController.GetFile (strContPath)
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open(sTestSuiteControllerFileSpec.Path)
	intRow = 2	
	Do Until objExcel.Cells(intRow,1).Value = ""
		Set sAppController 	= objWorkbook.Sheets("Controller")
		strTestname  = sAppController.Cells(intRow, 3).Value
		If strTestname = Environment.Value ("TestName") Then
			strBrowser  = UCASE(sAppController.Cells(intRow, 2).Value)
                 If INSTR(1,UCASE(strBrowser),"IE") > 0 Then
					 strBrowser = "IE"
				 End If
				 If INSTR(1,UCASE(strBrowser),"CHROME") > 0 Then
					 strBrowser = "CHROME"
				 End If
				 If INSTR(1,UCASE(strBrowser),"FIREFOX") > 0 Then
					 strBrowser = "FIREFOX"
				 End If
			Exit Do
		End If
		intRow = intRow +1 
	Loop
	objWorkbook.Close '  Close the excel report
	objExcel.Quit
	Set objExcel = Nothing
	Set objWorkbook = Nothing
	Set fsoTestSuiteController = nothing
	Set sTestSuiteControllerFileSpec = nothing
	If strBrowser = "" Then
		strBrowser = "CHROME"
	End If
	systemutil.CloseProcessByName("EXCEL.EXE")
	'-----------------------------------------------------'
	'---------------Find The Test URL---------------------'
	'-----------------------------------------------------'
	
	Set objExcel = CreateObject("Excel.Application") 
	objExcel.Visible = True 
	Set controllerExcel = objExcel.Workbooks.Open( strContPath ) 
	Set controllerObj = controllerExcel.Sheets("Controller")
	rowCountController = controllerObj.UsedRange.Rows.Count
	Set environmentObj = controllerExcel.Sheets("Environment")
	rowCountEnvironment = environmentObj.UsedRange.Rows.Count
	testName = Environment.Value("TestName")
	For contRow = 2 to rowCountController
		If testName =  controllerObj.cells(contRow ,3).value Then
			env = controllerObj.cells(contRow,7).value
			For envRow = 2 to rowCountEnvironment
				If env = environmentObj.cells(envRow, 1).value Then
					strURL = environmentObj.cells(envRow, 2).value
				End If
			Next
			Exit For
		End If
	Next
	controllerExcel.close
	'objExcel.Quit
	Set objExcel = nothing
	Set controllerExcel = nothing
	Set controllerObj = nothing
	Set environmentObj = nothing
	'systemutil.CloseProcessByName("EXCEL.EXE")
	'-----------------------------------------------------'

	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}" & "!\\.\root\cimv2")
	Set colProcess = objWMIService.ExecQuery ("Select * From Win32_Process")
	For Each objProcess in colProcess
		If LCase(objProcess.Name) = LCase("EXCEL.EXE") OR LCase(objProcess.Name) = LCase("EXCEL.EXE *32") Then
	        objProcess.Terminate()
	        ' MsgBox "- ACTION: " & objProcess.Name & " terminated"
		End If
	Next
	For Each objProcess in colProcess
		If LCase(objProcess.Name) = LCase("WerFault.exe") OR LCase(objProcess.Name) = LCase("WerFault.exe *32") Then
        objProcess.Terminate()
        ' MsgBox "- ACTION: " & objProcess.Name & " terminated"
		End If
	Next

End Function



' ================================================================================================
'  NAME				: booleanDBConnection
'  DESCRIPTION 	  	: This function returns a boolean True If SQL Server Connection is successful
'  PARAMETERS		: Connection Object , String server , String databse
' ================================================================================================
Public Function booleanDBConnection(ByRef oConn , strServer , strDatabase)
	Set oDataBase = getExcelDataObject("DataBase")
	'NOTE : ID PASS not needed as we are using Windows Authentication
	strCon = "Driver={SQL Server};Server="& strServer &";database="& strDatabase
	set oConn = CreateObject("ADODB.Connection")
	oConn.Open strCon
	Set oDataBase = Nothing
	If Err.Number <> 0 Then
		logResult micFail , "DB Connection Should be Established" , "Failed to connect with DB"
		booleanDBConnection = False
	Else
		logResult micPass , "DB Connection Should be Established" , "DB Connection Established"
		booleanDBConnection = True
	End If
End Function


' ================================================================================================
'  NAME				: booleanDBQuery
'  DESCRIPTION 	  	: This function sets the result set of the executed query
'  PARAMETERS		: Query string , Connection Objec , Result Set object
' ================================================================================================
Public Function booleanDBQuery(strQuery , oConn , oRst)
Dim oField 
set oRst = CreateObject("ADODB.recordset")
oRst.Open strQuery, oConn
If Err.Number <> 0 Then
	logResult micFail , "Fetching Data from DB" , "Data Fetching Failed"
	booleanDBQuery = False
Else
	logResult micPass , "Fetching Data from DB" , "Data Fetched Successfully"
	booleanDBQuery = True
End If
End Function




' ================================================================================================
'  NAME				: booleanDBDisconnect
'  DESCRIPTION 	  	: This function returns true when Disconnection is successful
'  PARAMETERS		: result set object, connection object
' ================================================================================================
Public Function booleanDBDisconnect(oRst , oConn)
oRst.close
oConn.close
If Err.Number <> 0 Then
	logResult micFail , "DB Connection Should be Closed" , "DB Connection Failed to close"
	booleanDBDisconnect = False
Else
	logResult micPass , "DB Connection Should be Closed" , "DB Connection closed"
	booleanDBDisconnect = True
End If
End Function
