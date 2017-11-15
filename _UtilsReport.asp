<%
'list of functions used specific to language bank
Function GetLangSurvey(lngID)
	If lngID = 3 Then 
		GetLangSurvey = "clientsurveyArabic.pdf"
	ElseIf lngID = 10 Then
		GetLangSurvey = "clientsurveyFarsi.pdf"
	ElseIf lngID = 17 Then
		GetLangSurvey = "clientsurveyKorean.pdf"
	ElseIf lngID = 49 Then
		GetLangSurvey = "clientsurveyNepali.pdf"
	ElseIf lngID = 21 Then
		GetLangSurvey = "clientsurveyPortuguese.pdf"
	ElseIf lngID = 22 Then
		GetLangSurvey = "clientsurveyRussian.pdf"
	ElseIf lngID = 24 Then
		GetLangSurvey = "clientsurveySomali.pdf"
	ElseIf lngID = 25 Then
		GetLangSurvey = "clientsurveySpanish.pdf"
	ElseIf lngID = 29 Then
		GetLangSurvey = "clientsurveyVietnamese.pdf"
	Else
		GetLangSurvey = "clientsurveyEnglish.pdf"
	End If
End Function
Function Z_GetClass(deptID)
	Z_GetClass = ""
	Set rsClass = Server.CreateObject("ADODB.RecordSet")
	rsClass.Open "SELECT [class] FROM dept_T WHERE [index] = " & deptID, g_strCONN, 3, 1
	If Not rsClass.EOF Then
		Z_GetClass = rsClass("class")
	End If
	rsClass.Close
	Set rsClass = Nothing
End Function
Function Z_GetInfoFROMAppID(AppID, infoneeded)
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	rsIntr.Open "SELECT " & infoneeded & " FROM request_T WHERE [index] = " & AppID, g_strCONN, 3, 1
	If Not rsIntr.EOF Then
		Z_GetInfoFROMAppID = rsIntr(infoneeded)
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End Function
Function AppAssigned(appId)
	AppAssigned = False
	Set rsReq = Server.CreateObject("ADODB.RecordSet")
	rsReq.Open "SELECT IntrID FROM request_T WHERE [index] = " & appID, g_strCONN, 3, 1
	If Z_CZero(rsReq("IntrID")) > 0 Then AppAssigned = True
	rsReq.Close
	Set rsReq = Nothing
End Function
Function SaveHist(xxx, mypage)
SaveHist = false
	'SAVE HIST SQL
	server.scripttimeout = 360000
	tmpHist = ""
	Set rsHist = Server.CreateObject("ADODB.RecordSet")
	Set rsLB = Server.CreateObject("ADODB.RecordSet")
	sqlHist = "SELECT * FROM hist_T WHERE [timestamp] = '" & Now & "'"
	sqlLB = "SELECT * FROM request_T WHERE [index] = " & xxx
	rsLB.Open sqlLB, g_strCONN, 1, 3
	rsHist.Open sqlHist, g_strCONNHist2, 1,3 
	If not rsLB.EOF Then
		rsHist.AddNew
		rsHist("LBID") = xxx
		rsHist("Timestamp") = Now
		rsHist("Author") = Request.Cookies("LBUsrName")
		rsHist("pageused") = mypage
		x = 1
On error resume next
    Do Until x = rsLB.Fields.Count
    	If x = 7 Then 
    		tmpHist = tmpHist & """" & rsLB.Fields(x).Value & "|" & GetLang(rsLB.Fields(x).Value) & ""","
    	ElseIf x = 19 Then
    		tmpHist = tmpHist & """" & rsLB.Fields(x).Value & "|" & GetInst(rsLB.Fields(x).Value) & ""","
    	ElseIf x = 20 Then
    		tmpHist = tmpHist & """" & rsLB.Fields(x).Value & "|" & GetDept(rsLB.Fields(x).Value) & ""","
    	ElseIf x = 23 Then
    		tmpHist = tmpHist & """" & rsLB.Fields(x).Value & "|" & GetIntr(rsLB.Fields(x).Value) & ""","
    	Else
        tmpHist = tmpHist & """" & rsLB.Fields(x).Value & ""","
      End If
        x = x + 1
    Loop
    rsHist("Hist") = trim(tmpHist)
		rsHist.Update
	End If
	rsLB.CLose
	set rsLB = Nothing
	rsHist.Close
	Set rsHist = Nothing
	SaveHist = True
End Function
Function GetAppDate(xxx)
	Set rsLang = Server.CreateObject("ADODB.RecordSet")
	sqlLang = "SELECT CONVERT(varchar(10), appdate, 101) as myAppDate FROM request_T WHERE [index] = " & xxx
	rsLang.Open sqlLang, g_strCONN, 3, 1
	If Not rsLang.EOF Then
		GetAppDate = rsLang("myAppDate")
	End If
	rsLang.Close
	Set rsLang = Nothing
End Function
Function SearchArraysHours(myIntr, tmpIntr)
	DIM	lngMax, lngI
	SearchArraysHours = -1
	On Error Resume Next	
	lngMax = UBound(tmpIntr)
	If Err.Number <> 0 Then Exit Function
	For lngI = 0 to lngMax
		If tmpIntr(lngI) = myIntr Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArraysHours = lngI
End Function
Function GetFileNum(xxx)
	GetFileNum = ""
	Set rsCity = Server.CreateObject("ADODB.RecordSet")
	sqlCity = "SELECT FileNum FROM interpreter_T WHERE [index] = " & xxx
	rsCity.Open sqlCity, g_strCONN,1, 3
	If Not rsCity.EOF Then
		GetFileNum = rsCity("FileNum")
	End If
	rsCity.Close
	Set rsCity = Nothing
End Function
Function GetCity(xxx)
	GetCity = ""
	Set rsCity = Server.CreateObject("ADODB.RecordSet")
	sqlCity = "SELECT City FROM dept_T WHERE [index] = " & xxx
	rsCity.Open sqlCity, g_strCONN,1, 3
	If Not rsCity.EOF Then
		GetCity = rsCity("City")
	End If
	rsCity.Close
	Set rsCity = Nothing
End Function
Function CleanFax(strFax)
	CleanFax = Replace(strFax, "-", "") 
End Function
Function IsActive(xxx)
	'check if interpreter is active
	IsActive = True
	Set rsAct = Server.CreateObject("ADODB.RecordSet")
	sqlAct = "SELECT Active FROM interpreter_T WHERE [index] = " & xxx
	rsAct.Open sqlAct, g_strCONN, 3, 1
	If Not rsAct.EOF Then
		If rsAct("Active") = False Then IsActive = False
	End If
	rsAct.Close
	Set rsAct = Nothing	
End Function
Function GetReas(xxx)
	If xxx = "" THen
		GetReas = ""
		exit function
	End IF
	GetReas = ""
	tmpReas = Split(xxx, "|")
	CtrReas = Ubound(tmpReas)
	x = 0
	Do Until x = CtrReas + 1
		Set rsReas = Server.CreateObject("ADODB.RecordSet")
		sqlReas = "SELECT reason FROM Reason_T WHERE [index] = " & tmpReas(x)
		rsReas.Open sqlReas, g_strCONNHP, 3, 1
		If Not rsReas.EOF Then
			GetReas = GetReas & rsReas("reason") & "<br>"
		End If
		rsReas.Close
		Set rsReas = Nothing
		x = x + 1
	Loop
End Function
'GET STATUS
Function GetStat(zzz)
	Select Case zzz
		Case 0 GetStat = "Pending"
		Case 1 GetStat = "Completed"
		Case 2 GetStat = "Missed"
		Case 3 GetStat = "Canceled"
		Case 4 GetStat = "Canceled-Billable"
	End Select
End Function
'GET LANGUAGE
Function GetLang(zzz)
	GetLang = "N/A"
	Set rsLang = Server.CreateObject("ADODB.RecordSet")
	sqlLang = "SELECT [Language] FROM language_T WHERE [index] = " & zzz
	rsLang.Open sqlLang, g_strCONN, 3, 1
	If Not rsLang.EOF Then
		GetLang = rsLang("Language")
	End If
	rsLang.Close
	Set rsLang = Nothing
End Function
'GET INSTITUTION w/ CLASSIFICATION
Function GetInst(zzz)
	GetInst = "N/A"
	If zzz = "" Then Exit Function
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT Facility FROM institution_T WHERE [index] = " & zzz
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		GetInst = rsInst("Facility")
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
'GET CLASS
Function GetClass(zzz)
GEtcLAss = zzz
if Z_CZero(zzz) = 0 then exit function
	Select Case zzz
		Case 1 GetClass = "Social Services"
		Case 2 GetClass = "Private"
		Case 3 GetClass = "Court"
		Case 4 GetClass = "Medical"
		Case 5 GetClass = "Legal"
		Case 6 GetClass = "Mental Health"
	End Select
End Function
'GET INTERPRETER
Function GetIntr(zzz)
	GetIntr = "N/A"
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT [Last Name], [First Name] FROM interpreter_T WHERE [index] = " & zzz
	rsIntr.Open sqlIntr, g_strCONN, 3, 1
	If Not rsIntr.EOF Then
		GetIntr = rsIntr("Last Name") & ", " & rsIntr("First Name")
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End Function
'GET INTERPRETER
Function GetIntr2(zzz)
	GetIntr2 = "N/A"
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT [Last Name], [First Name] FROM interpreter_T WHERE [index] = " & zzz
	rsIntr.Open sqlIntr, g_strCONN, 3, 1
	If Not rsIntr.EOF Then
		GetIntr2 = rsIntr("First Name") & " " & rsIntr("Last Name")
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End Function
'SEARCH FOR TOWNS
Function SearchArraysTown(xtown, xname, xlang, xclass, tmpTown, tmpName, tmpLang, tmpClass)
	DIM	lngMax, lngI
	SearchArraysTown = -1
	On Error Resume Next	
	lngMax = UBound(tmpTown)
	If Err.Number <> 0 Then Exit Function
	For lngI = 0 to lngMax
		If tmpTown(lngI) = xtown And tmpName(lngI) = xname And tmpLang(lngI) = xlang And tmpClass(lngI) = xclass Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArraysTown = lngI
End Function
'SEARCH FOR INSTITUTION
Function SearchArraysInst(xinst, xdept, xname, xlang, xclass, tmpInst, tmpDept, tmpName, tmpLang, tmpClass)
	DIM	lngMax, lngI
	SearchArraysInst = -1
	On Error Resume Next	
	lngMax = UBound(tmpInst)
	If Err.Number <> 0 Then Exit Function
	For lngI = 0 to lngMax
		If tmpInst(lngI) = xinst And tmpDept(lngI) = xdept And tmpName(lngI) = xname And tmpLang(lngI) = xlang And tmpClass(lngI) = xclass Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArraysInst = lngI
End Function
'SEARCH FOR INSTITUTION 2
Function SearchArraysInst2(xinst, xdept, xlang, tmpInst, tmpDept, tmpLang)
	DIM	lngMax, lngI
	SearchArraysInst2 = -1
	On Error Resume Next	
	lngMax = UBound(tmpInst)
	If Err.Number <> 0 Then Exit Function
	For lngI = 0 to lngMax
		If tmpInst(lngI) = xinst And tmpDept(lngI) = xdept And tmpLang(lngI) = xlang Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArraysInst2 = lngI
End Function
'SEARCH FOR INTERPRETER
Function SearchArraysIntr(xname, xinst, xlang, xclass, tmpIntrName, tmpInst, tmpLang, tmpClass)
	DIM	lngMax, lngI
	SearchArraysIntr = -1
	On Error Resume Next	
	lngMax = UBound(tmpIntrName)
	If Err.Number <> 0 Then Exit Function
	For lngI = 0 to lngMax
		If tmpIntrName(lngI) = xname And tmpInst(lngI) = xinst And tmpLang(lngI) = xlang And tmpClass(lngI) = xclass Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArraysIntr = lngI
End Function
Function CheckApp(tmpdate)
	CheckApp = "#FFFFFF"
	If Request.Cookies("LBUSERTYPE") <> 2 Then
		'sqlReq = "SELECT appDate FROM request_T WHERE appDate = #" & tmpDate & "# ORDER BY appTimeFrom"
	Else
		Set rsReq = Server.CreateObject("ADODB.RecordSet")
		sqlReq = "SELECT appDate FROM request_T WHERE appDate = '" & tmpDate & "' AND IntrID = " & Session("UIntr") & " " & _
			"AND showintr = 1 AND NOT(STATUS = 2 OR STATUS = 3) ORDER BY appTimeFrom"
		rsReq.Open sqlReq, g_strCONN, 3, 1
		If Not rsReq.EOF Then
			CheckApp = "#FFFFCE"
		End If
		rsReq.Close
		Set rsReq = Nothing
	End If
	
End Function
Function GetReq(zzz)
	GetReq = "N/A"
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT Lname, Fname FROM requester_T WHERE [index] = " & zzz
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		GetReq = rsRP("Lname") & ", " & rsRP("Fname")
	End If
	rsRP.Close
	Set rsRP = Nothing
End Function
'GET INSTITUTION's NAME
Function GetInst2(zzz)
	GetInst2 = "N/A"
	If zzz = "" Then Exit Function
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT Facility FROM institution_T WHERE [index] = " & zzz
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		tmpIname = rsInst("Facility") 
		GetInst2 = tmpIname
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
'GET INSTITUTION's ADDRESS
Function GetInst3(zzz)
	GetInst3 = "N/A"
	If zzz = "" Then Exit Function
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT Address, City, State, Zip FROM institution_T WHERE [index] = " & zzz
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		GetInst3 = rsInst("Address") & ", "& rsInst("City") & ", " & rsInst("State") & ", " & rsInst("Zip")
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
'GET INSTITUTION's NAME
Function GetInst4(zzz)
	GetInst4 = "N/A"
	If zzz = "" Then Exit Function
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT Facility, Department FROM institution_T WHERE [index] = " & zzz
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		tmpIname = """" & rsInst("Facility") & """"
		If rsInst("Department") <> "" Then tmpIname = """" & rsInst("Facility") & " - " & rsInst("Department") & """"
		GetInst4 = tmpIname
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
'GET INSTITUTION's ADDRESS for CSV
Function GetInst5(zzz)
	GetInst5 = "N/A"
	If zzz = "" Then Exit Function
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT Address, City, State, Zip FROM institution_T WHERE [index] = " & zzz
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		GetInst5 = """" & rsInst("Address") & ""","& rsInst("City") & "," & rsInst("State") & "," & rsInst("Zip")
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
Function GetDept(xxx)
	GetDept = ""
	If xxx = "" Then Exit Function
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT Dept FROM dept_T WHERE [index] = " & xxx
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		GetDept = rsInst("Dept")
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
Function GetMyDept(xxx)
	GetMyDept = ""
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	sqlDept = " SELECT Dept FROM dept_T WHERE [index] = " & xxx
	rsDept.Open sqlDept, g_strCONN, 3, 1
	If Not rsDept.EOF Then
		GetMyDept = " - " & rsDept("Dept")
	End If
	rsDept.Close
	Set rsDept = Nothing
End Function
Function GetDeptAdr(xxx)
	GetDeptAdr = ""
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	sqlDept = " SELECT Address, City, State, Zip FROM dept_T WHERE [index] = " & xxx
	rsDept.Open sqlDept, g_strCONN, 3, 1
	If Not rsDept.EOF Then
		GetDeptAdr = rsDept("Address") & ", " & rsDept("City") & ", " & rsDept("State") & ", " & rsDept("Zip")
	End If
	rsDept.Close
	Set rsDept = Nothing
End Function
Function GetInstDept(xxx)
	GetInstDept = ""
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	sqlDept = " SELECT InstID FROM dept_T WHERE [index] = " & xxx
	rsDept.Open sqlDept, g_strCONN, 3, 1
	If Not rsDept.EOF Then
		GetInstDept = GetInst2(rsDept("InstID"))
	End If
	rsDept.Close
	Set rsDept = Nothing
End Function
'SEARCH STATS
Function SearchStats(xFac, xMonthYr, tmpFac, tmpMonthYr)
	DIM lngMax, lngI
	SearchStats = -1
	On Error Resume Next
	lngMax = UBound(tmpFac)
	If Err.Number <> 0 Then Exit Function
	For lngI = 0 to lngMax
		If tmpFac(lngI) = xFac And tmpMonthYr(lngI) = xMonthYr Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchStats = lngI
End Function
Function GetMisReason(xxx)
	GetMisReason = "N/A"
	Set rsMis = Server.CreateObject("ADODB.RecordSet")
	sqlMis = "SELECT reason FROM missed_T WHERE [index] = " & xxx
	rsMis.Open sqlMis, g_strCONN, 3, 1
	If Not rsMis.EOF Then
		GetMisReason = rsMis("reason")
	End If
	rsMis.Close
	Set rsMis = Nothing
End Function
Function GetCanReason(xxx)
	GetCanReason = "N/A"
	Set rsMis = Server.CreateObject("ADODB.RecordSet")
	sqlMis = "SELECT reason FROM cancel_T WHERE [index] = " & xxx
	rsMis.Open sqlMis, g_strCONN, 3, 1
	If Not rsMis.EOF Then
		GetCanReason = rsMis("reason")
	End If
	rsMis.Close
	Set rsMis = Nothing
End Function
Function Z_FormatTime(xxx, zzz)
	Z_FormatTime = ""
	If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, zzz)
End Function
%>