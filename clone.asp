<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Security.asp" -->
<%
If Cint(Request.Cookies("LBUSERTYPE")) = 2 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
Function CleanMe(xxx)
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, "'", " ")
End Function
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
TimeNow = Now
If Request("Clone") <> "" Then
	Set rsClone = Server.CreateObject("ADODB.RecordSet")
	sqlClone = "SELECT * FROM request_T WHERE [index] = " & Request("Clone")
	rsClone.Open sqlClone, g_strCONN, 3, 1
	If Not rsCLone.EOF Then
		tmpReqP = rsClone("ReqID") 
		tmplName = rsClone("clname") 
		tmpfName = rsClone("cfname")	
		chkClient = 0
		If rsClone("Client") = True Then chkClient = 1
		chkUClientadd = 0
		If  rsClone("CliAdd") = True Then chkUClientadd = 1
		tmpAddr = rsClone("caddress") 
		tmpCity = rsClone("cCity") 
		tmpState = rsClone("cstate") 
		tmpZip = rsClone("czip")
		tmpCAdrI = rsClone("CliAdrI")
		tmpCFon = rsClone("Cphone")
		tmpCAFon = rsClone("CAphone")
		tmpDir = rsClone("directions")
		tmpSC = rsClone("spec_cir")
		tmpDOB = rsClone("DOB")
		tmpLang = rsClone("langID")
		tmpAppDate = rsClone("appDate")
		tmpAppTFrom = Z_FormatTime(rsClone("appTimeFrom"))
		tmpAppTTo = Z_FormatTime(rsClone("appTimeTo"))
		tmpAppLoc = rsClone("appLoc")
		tmpInst = rsClone("instID")
		tmpDept = rsClone("deptID")
		tmpInstRate = rsClone("InstRate")
		tmpDoc = rsClone("docNum")
		tmpCRN = rsClone("CrtRumNum")
		tmpIntr = rsClone("IntrID")
		tmpIntrRate = rsClone("IntrRate")
		tmpEmer = 0
		If rsClone("Emergency") = True Then tmpEmer = 1
		tmpEmerFee = 0
		If rsClone("EmerFee") = True Then tmpEmerFEE = 1
		If rsClone("Gender") = vbNull Then
			tmpGender = -1
		Else
			tmpGender	= Z_CZero(rsClone("Gender"))
		End If
		'tmpMale = ""
		'tmpFemale = ""
		'If tmpGender = 0 Then 
		'	tmpMale = "SELECTED"
		'Else
		'	tmpFemale = "SELECTED"
		'End If
		chkMinor = 0
		If rsClone("Child") Then chkMinor = 1
		'tmpHPID = rsClone("HPID")
	End If
	rsClone.CLose
	Set rsClone = Nothing
	'SAVE IN DB
	Set rsMain = Server.CreateObject("ADODB.RecordSet")
	sqlApp = "SELECT * FROM Request_T"
	rsMain.Open sqlApp, g_strCONN, 1, 3
	rsMain.AddNew
	rsMain("timestamp") = TimeNow
	rsMain("reqID") = tmpReqP
	rsMain("clname") = tmplName
	rsMain("cfname") = tmpfName
	rsMain("Caddress") = tmpAddr
	rsMain("Ccity") = tmpCity
	rsMain("Cstate") = tmpState
	rsMain("Czip") = tmpZip
	rsMain("directions") = tmpDir
	rsMain("spec_cir") = tmpSC
	rsMain("DOB") = Z_DateNull(tmpDOB)
	rsMain("LangID") = tmpLang
	rsMain("appDate") = date
	rsMain("appTimeFrom") = date & " " & tmpAppTFrom
	rsMain("appTimeTo") = date & " " & tmpAppTTo
	rsMain("appLoc") = tmpAppLoc
	rsMain("InstID") = tmpInst
	rsMain("DeptID") = tmpDept
	rsMain("InstRate") = tmpInstRate
	rsMain("docNum") = tmpDoc
	rsMain("CrtRumNum") = tmpCRN
	If chkClient = 1 Then rsMain("Client") = True
	rsMain("Cphone") = tmpCFon
	rsMain("IntrID") = tmpIntr
	rsMain("IntrRate") = Z_CDbl(tmpIntrRate)
	If tmpEmer = 1 Then rsMain("Emergency") = True
	'rsMain("Comment") = tmpEntry(25)
	rsMain("CAphone") = tmpCAFon
	If chkUClientadd = 1 Then rsMain("CliAdd") = True
	rsMain("CliAdrI") = CleanMe(tmpCAdrI)
	'rsMain("IntrComment") = tmpEntry(31)
	If tmpEmerFEE = 1 Then rsMain("EmerFee") = true
	'rsMain("BilComment") = tmpEntry(33)
	'rsMain("LBcomment") = tmpEntry(34)
	'response.write "<!---" & tmpEntry(35) & "-->"
	If tmpGender < 0 Then
		rsMain("Gender") = vbNull
	Else
		rsMain("Gender") = tmpGender
	End If
	rsMain("Child") = false
	If chkMinor = 1 Then rsMain("Child") = true
	rsMain.Update
	tmpID = rsMain("index")
	rsMain.Close
	Set rsMain = Nothing
	'SAVE HISTORY

	Set rsHist = Server.CreateObject("ADODB.RecordSet")
	sqlHist = "SELECT * FROM History_T"
	rsHist.Open sqlHist, g_strCONNHist, 1,3 
	rsHist.AddNew
	rsHist("reqID") = tmpID
	rsHist("Creator") = Request.Cookies("LBUsrName")
	rsHist("date") = date
	rsHist("dateTS") = TimeNow
	rsHist("dateU") = Request.Cookies("LBUsrName")
	rsHist("Stime") = tmpAppTFrom
	rsHist("StimeTS") = TimeNow
	rsHist("StimeU") = Request.Cookies("LBUsrName")
	If chkClient = 1 Then
		tmpHistAdr = tmpAddr & "|" & tmpCity & "|" & tmpState & "|" & tmpZip
	Else
		tmpHistAdr = "Department Address"
	End If
	rsHist("location") = tmpHistAdr
	rsHist("locationTS") = TimeNow
	rsHist("locationU") = Request.Cookies("LBUsrName")
	If tmpIntr <> "-1" Or tmpIntr = 0 Then
		rsHist("interID") = tmpIntr
		rsHist("interTS") = TimeNow
		rsHist("interU") = Request.Cookies("LBUsrName")
	End If
	rsHist.Update
	rsHist.Close
	Set rsHist = Nothing
	
	Session("MSG") = "NOTE: Entries cloned from Request: " & Request("Clone")
	Response.Redirect "reqconfirm.asp?ID=" & tmpID
End If
%>