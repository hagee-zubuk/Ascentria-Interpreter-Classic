<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) = 2 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
Function GetMyStatus(xxx)
	Select Case xxx
		Case 1
			GetMyStatus = "COMPLETED"
		Case 2
			GetMyStatus = "MISSED"
		Case 3
			GetMyStatus = "CANCELED"
		Case 4
			GetMyStatus = "CANCELED (BILLABLE)"
		Case Else
			GetMyStatus = "PENDING"
	End Select
End Function
Function CleanMe(xxx)
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, "'", " ")
End Function
Function CleanFax(strFax)
	CleanFax = Replace(strFax, "-", "") 
End Function
Function GetPrime(xxx)
	'get primary number of requesting person
	GetPrime = ""
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT * FROM requester_T WHERE [index] = " & xxx
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		If rsRP("prime") = 0 Then
			GetPrime = rsRP("Email")
		ElseIf rsRP("prime") = 1 Then
			GetPrime = ""
		ElseIf rsRP("prime") = 2 Then
			GetPrime = CleanFax(Trim(rsRP("Fax"))) & "@emailfaxservice.com" 
		End If
	End If
	rsRP.Close
	set rsRP = Nothing
End Function
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
server.scripttimeout = 360000
tmpPage = "document.frmAssign."
'GET REQUEST
Set rsConfirm = Server.CreateObject("ADODB.RecordSet")
sqlConfirm = "SELECT * FROM Request_T WHERE [index] = " & Request("ID")
rsConfirm.Open sqlConfirm, g_strCONN, 3, 1
If Not rsConfirm.EOF Then
	TS = rsConfirm("timestamp")
	RP = rsConfirm("reqID") 
	tmpClient = ""
	tmpDeptaddr = ""
	tmplName = rsConfirm("clname") 
	tmpfName = rsConfirm("cfname")
	chkClient = ""
	If rsConfirm("Client") = True Then chkClient = "checked"
	chkUClientadd = ""
	If  rsConfirm("CliAdd")  = True Then chkUClientadd = "checked"
	tmpAddr = rsConfirm("caddress") 
	tmpCity = rsConfirm("cCity") 
	tmpState = rsConfirm("cstate") 
	tmpZip = rsConfirm("czip")
	tmpCAdrI = rsConfirm("CliAdrI")
	tmpCFon = rsConfirm("Cphone")
	tmpCAFon = rsConfirm("CAphone")
	tmpDir = rsConfirm("directions")
	tmpSC = rsConfirm("spec_cir")
	tmpDOB = rsConfirm("DOB")
	tmpLang = rsConfirm("langID")
	tmpAppDate = rsConfirm("appDate")
	tmpAppTFrom = Z_FormatTime(rsConfirm("appTimeFrom"))
	tmpAppTTo = Z_FormatTime(rsConfirm("appTimeTo"))
	tmpAppLoc = rsConfirm("appLoc")
	tmpInst = rsConfirm("instID")
	tmpDept = rsConfirm("deptID")
	tmpInstRate = rsConfirm("InstRate")
	tmpDoc = rsConfirm("docNum")
	tmpCRN = rsConfirm("CrtRumNum")
	tmpInst = rsConfirm("instID")
	tmpDept = rsConfirm("DeptID")
	tmpInstRate = Z_FormatNumber(rsConfirm("InstRate"), 2)
	tmpDoc = rsConfirm("docNum")
	tmpCRN = rsConfirm("CrtRumNum")
	tmpIntr = rsConfirm("IntrID")
	tmpIntrRate = Z_FormatNumber(rsConfirm("IntrRate"), 2)
	tmpEmer = ""
	If rsConfirm("Emergency") = True Then tmpEmer = "(EMERGENCY)" 
	tmpCom = rsConfirm("Comment")
	tmpComintr = rsConfirm("IntrComment")
	tmpcombil = rsConfirm("bilComment")
	Statko = GetMyStatus(rsConfirm("Status"))
	tmpBilHrs = rsConfirm("Billable")
	tmpActTFrom = Z_FormatTime(rsConfirm("astarttime")) 
	tmpActTTo = Z_FormatTime(rsConfirm("aendtime"))
	tmpBilTInst = rsConfirm("TT_Inst")
	tmpBilTIntr = rsConfirm("TT_Intr")
	tmpBilMInst = rsConfirm("M_Inst")
	tmpBilMIntr = rsConfirm("M_Intr")
	tmpGender	= Z_CZero(rsConfirm("Gender"))
	tmpMale = ""
	tmpFemale = ""
	If tmpGender = 0 Then 
		tmpMale = "SELECTED"
	Else
		tmpFemale = "SELECTED"
	End If
	chkMinor = ""
	If rsConfirm("Child") Then chkMinor = "CHECKED"
	'timestamp on sent/print
	tmpSentReq = "Request email has not been sent to Requesting Person."
	If rsConfirm("SentReq") <> "" Then tmpSentReq = "Request email was last sent to Requesting Person on <b>" & rsConfirm("SentReq") & "</b>."
	tmpSentIntr = "Request email has not been sent to Interpreter."
	If rsConfirm("SentIntr") <> "" Then tmpSentIntr = "Request email was last sent to Interpreter on <b>" & rsConfirm("SentIntr") & "</b>."
	tmpPrint = "Request has not been printed."
	If rsConfirm("Print") <> "" Then tmpPrint = "Request was last printed on<b> " & rsConfirm("Print") & "</b>."
	tmpHPID = Z_CZero(rsConfirm("HPID"))
	tmpLBcom = rsConfirm("LBcomment")
End If
rsConfirm.Close
Set rsConfirm = Nothing
'RESUPPLY DATA IN ERROR EVENT
If Session("MSG") <> "" And Request("ID") = "" Then
	tmpEntry = Split(Z_DoDecrypt(Request.Cookies("LBREQUEST")), "|")
	tmpNewInst = Split(Z_DoDecrypt(Request.Cookies("LBINST")), "|")
	tmpNewDept = Split(Z_DoDecrypt(Request.Cookies("LBDEPT")), "|")
	tmpNewReq = Split(Z_DoDecrypt(Request.Cookies("LBREQ")), "|")
	tmpNewIntr = Split(Z_DoDecrypt(Request.Cookies("LBINTR")), "|")
	tmpNewIntrBTN = tmpNewIntr(10)
	If tmpNewReq(6) = "BACK" Then
		tmpNewReqLN = tmpNewReq(0)
		tmpNewReqFN = tmpNewReq(1)
		tmpReqExt = tmpNewReq(8)
		tmpNewReqPhone = tmpNewReq(2)
		tmpNewReqEmail = tmpNewReq(3)
		tmpNewReqFax = tmpNewReq(4)
		tmpNewReqPrim = tmpNewReq(7)
		selRPEmail = ""
		selRPPhone = ""
		selRPFax = ""
		if tmpNewReqPrim = 0 Then selRPEmail = "checked"
		if tmpNewReqPrim = 1 Then selRPPhone = "checked"
		if tmpNewReqPrim = 2 Then selRPFax = "checked"
	Else
		tmpReqP = tmpEntry(1)
	End If
	tmplName = tmpEntry(2)
	tmpfName = tmpEntry(3)
	chkClient = ""
	If tmpEntry(20) <> "" Then chkClient = "checked"
	tmpAddr = tmpEntry(4)
	chkUClientadd = ""
	If tmpEntry(29) <> "" Then chkUClientadd = "checked"
	tmpCity = tmpEntry(5)
	tmpState = tmpEntry(6)
	tmpZip = tmpEntry(7)
	tmpCAdrI = tmpEntry(30)
	tmpDir = tmpEntry(8)
	tmpSC = tmpEntry(9)
	tmpDOB = tmpEntry(10)
	tmpLang = tmpEntry(11)
	tmpAppDate = tmpEntry(12)
	tmpAppTFrom = Z_FormatTime(tmpEntry(13))
	tmpAppTTo = Z_FormatTime(tmpEntry(14))
	tmpAppLoc = tmpEntry(15)
	If tmpNewInst(6) = "BACK" Then 
		tmpNewInstchk = tmpNewInst(6)
		tmpNewInstTxt = tmpNewInst(0)
	Else
		tmpInst = tmpEntry(16)
	End If
	If tmpNewDept(6) = "BACK" Then
		tmpNewInstchk = tmpNewDept(6)
		tmpNewInstDept = tmpNewDept(0)
		SocSer = ""
		Priv = ""
		Legal = ""
		Med = ""
		tmpClass = tmpNewDept(7)
		If tmpNewDept(7) = 1 Then SocSer = "selected"
		If tmpNewDept(7) = 2 Then Priv = "selected"
		If tmpNewDept(7) = 3 Then Court = "selected"
		If tmpNewDept(7) = 4 Then Med = "selected"
		If tmpNewDept(7) = 5 Then Legal = "selected"
		tmpNewInstAddr = tmpNewDept(2)
		tmpNewInstCity = tmpNewDept(3)
		tmpNewInstState = tmpNewDept(4)
		tmpNewInstZip = tmpNewDept(5)
		tmpNewInstAddrI = tmpNewDept(16)
		If tmpNewInst(8) <> "" Then 
			chkBillMe = "checked"
		Else
			tmpBilInstAddr = tmpNewDept(9)
			tmpBilInstCity = tmpNewDept(10)
			tmpBilInstState = tmpNewDept(11)
			tmpBilInstZip = tmpNewDept(12)
		End If	
		tmpBLname =  tmpNewDept(13)
	Else
		tmpDept = tmpEntry(26)
	End If
	tmpRate = tmpEntry(17)
	tmpDoc = tmpEntry(18)
	tmpCRN = tmpEntry(19)
	tmpCFon = tmpEntry(21)
	If Request.Cookies("LBACTION") = 1 Then
		tmpCAFon = tmpEntry(27)
	Else
		tmpCAFon = tmpEntry(38)
	End If
	If tmpNewIntrBTN = "BACK" Then
		tmpIntrLname = tmpNewIntr(0)
		tmpIntrFname = tmpNewIntr(1)
		tmpIntrEmail = tmpNewIntr(2)
		tmpIntrP1 = tmpNewIntr(3)
		tmpIntrExt = tmpNewIntr(13)
		tmpIntrFax = tmpNewIntr(4)
		tmpIntrP2 = tmpNewIntr(5)
		tmpIntrAddr = tmpNewIntr(6)
		tmpIntrCity = tmpNewIntr(7)
		tmpIntrState = tmpNewIntr(8)
		tmpNewIntrAddrI = tmpNewIntr(16)
		tmpIntrZip = tmpNewIntr(9)
		tmpInHouse = ""
		If tmpNewIntr(11) <> "" Then tmpInHouse = "checked"
		tmpIntrPrim = tmpNewIntr(12)
		selIntrFax = ""
		selIntrP2 = ""
		selIntrP1 = ""
		selIntrEmail = ""
		if tmpIntrPrim = 0 Then selIntrEmail = "checked"
		if tmpIntrPrim = 1 Then selIntrP1 = "checked"
		if tmpIntrPrim = 2 Then selIntrP2 = "checked"
		if tmpIntrPrim = 0 Then selIntrFax = "checked"
		tmpIntrRate =  tmpNewIntr(14)
	Else
		tmpIntr = tmpEntry(22)
	End If
	tmpEmer = ""
	If tmpEntry(24) <> "" Then tmpEmer = "checked"
	If Request.Cookies("LBACTION") = 1 Then	
		tmpCom =  tmpEntry(25)
	Else
		tmpCom =  tmpEntry(30)
	End If
End If
'GET INTERPRETER INFO
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE [index] = " & tmpIntr
rsIntr.Open sqlIntr, g_strCONN, 3, 1
If Not rsIntr.EOF Then
	tmpInHouse = ""
	If rsIntr("InHouse") = True Then tmpInHouse = "(In-House)"
	tmpIntrName = rsIntr("Last Name") & ", " & rsIntr("First Name") & " " & tmpInHouse
Else
	tmpIntrName = "<i>To be assigned.</i>"
	tmpIntr = 0
End If
rsIntr.Close
Set rsIntr = Nothing
'HP DATA
If tmpHPID <> 0  THen
	Set rsHP = Server.CreateObject("ADODB.RecordSet")
		sqlHP = "SELECT * FROM Appointment_T WHERE [index] = " & tmpHPID
	rsHP.Open sqlHP, g_StrCONNHP, 3, 1
	If Not rsHP.EOF Then
		tmpCallMe = ""
		If rsHP("callme") = True Then tmpCallMe = "* Call patient to remind of appointment"
		if rsHP("reason") <> "" Then mytmpReas = Z_Replace(rsHP("reason"),", ", "|")
		tmpReason = GetReas(mytmpReas)
		tmpClin = rsHP("clinician")  
		tmpReqname = rsHP("reqName")  
		InHP = 0
		tmpMeet = ""
		If rsHP("mwhere") = 1 Then
			InHP = 1
			tmpMeet = UCase(GetLoc(rsHP("mlocation")))
			If tmpMeet = "OTHER" Then tmpMeet = rsHP("mother")
		End If
		tmpParents = ""
		If rsHP("parents") <> "" Then tmpParents = rsHP("parents") 
	End If
	rsHp.Close
	Set rsHp = Nothing
End If
'GET REQUESTING PERSON
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM requester_T WHERE [index] = " & RP
rsReq.Open sqlReq, g_strCONN, 3, 1
If Not rsReq.EOF Then
	tmpRP = rsReq("Lname") & ", " & rsReq("Fname") 
	Fon = rsReq("phone") 
	If rsReq("pExt") <> "" Then Fon = Fon & " ext. " & rsReq("pExt")
	Fax = rsReq("fax")
	email = rsReq("email")
	Pcon = GetPrime(RP)
End If
rsReq.Close
Set rsReq = Nothing
'GET INSTITUTION
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T WHERE [index] = " & tmpInst
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	tmpIname = rsInst("Facility") 
End If
rsInst.Close
Set rsInst = Nothing 
'GET DEPARTMENT
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE [index] = " & tmpDept
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
	tmpDname = rsDept("dept") 
	tmpDeptaddr = rsDept("address") & ", " & rsDept("InstAdrI") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	tmpBaddr = rsDept("Baddress") & ", " & rsDept("BCity") & ", " &  rsDept("Bstate") & ", " & rsDept("Bzip")
	tmpBContact = rsDept("Blname")
	tmpZipInst = ""
	If rsDept("zip") <> "" Then tmpZipInst = rsDept("zip")
	If tmpDeptaddrG = "" Then 
		tmpDeptaddrG = rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	End If
End If
rsDept.Close
Set rsDept = Nothing 
'GET AVAILABLE LANGUAGES
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang = "SELECT * FROM language_T ORDER BY [Language]"
rsLang.Open sqlLang, g_strCONN, 3, 1
Do Until rsLang.EOF
	tmpL = ""
	If tmpLang = "" Then tmpLang = -1
	If CInt(tmpLang) = rsLang("index") Then tmpL = "selected"
	strLang = strLang	& "<option " & tmpL & " value='" & rsLang("Index") & "'>" &  rsLang("language") & "</option>" & vbCrlf
	strLangChk = strLangChk & "if (xxx == """ & Trim(rsLang("Language")) & """){ " & vbCrLf & _
		"return " & rsLang("index") & ";}"
	rsLang.MoveNext
Loop
rsLang.Close
Set rsLang = Nothing
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Language Bank - Interpreter Request Form - Edit Appointment - <%=Request("ID")%></title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function bawal(tmpform)
		{
			var iChars = ",|\"\'";
			var tmp = "";
			for (var i = 0; i < tmpform.value.length; i++)
			 {
			  	if (iChars.indexOf(tmpform.value.charAt(i)) != -1)
			  	{
			  		alert ("This character is not allowed.");
			  		tmpform.value = tmp;
			  		return;
		  		}
			  	else
		  		{
		  			tmp = tmp + tmpform.value.charAt(i);
		  		}
		  	}
		}
		function maskMe(str,textbox,loc,delim)
		{
			var locs = loc.split(',');
			for (var i = 0; i <= locs.length; i++)
			{
				for (var k = 0; k <= str.length; k++)
				{
					 if (k == locs[i])
					 {
						if (str.substring(k, k+1) != delim)
					 	{
					 		str = str.substring(0,k) + delim + str.substring(k,str.length);
		     			}
					}
				}
		 	}
			textbox.value = str
		}
		function KillMe(xxx)
		{
			var ans = window.confirm("Delete Request? Click Cancel to stop.");
			if (ans){
				document.frmAssign.action = "action.asp?ctrl=9&ReqID=" + xxx;
				document.frmAssign.submit();
			}
		}
		function SaveAss(xxx)
		{
			if (document.frmAssign.chkClientAdd.checked == true)
			{
				if (document.frmAssign.txtCliAdd.value == "" || document.frmAssign.txtCliCity.value == "" || document.frmAssign.txtCliState.value == "" || document.frmAssign.txtCliZip.value == "")
				{
					alert("Please input client's full address.")
					return;
				}
			}
			if (document.frmAssign.myLang.value == document.frmAssign.selLang.value)
			{
				document.frmAssign.action = "action.asp?ctrl=13&ReqID=" + xxx;
				document.frmAssign.submit();
			}
			else
			{
				var ans = window.confirm("Changing Language will result in assigned Interpreter being removed.\nClick Cancel to stop.");
				if (ans)
				{
					document.frmAssign.action = "action.asp?ctrl=13&Intr=1&ReqID=" + xxx;
					document.frmAssign.submit();
				}
			}
		}
		-->
		</script>
		<body onload=''>
			<form method='post' name='frmAssign'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td valign='top'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top' >
							<table cellSpacing='0' cellPadding='0' width="100%" border='0'>
									<!-- #include file="_greetme.asp" -->
								<tr>
								<td class='title' colspan='10' align='center'><nobr> Interpreter Request Form - Edit Appointment</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td  align='center' colspan='10'>
										<div name="dErr" style="width: 250px; height:55px;OVERFLOW: auto;">
											<table border='0' cellspacing='1'>		
												<tr>
													<td><span class='error'><%=Session("MSG")%></span></td>
												</tr>
											</table>
										</div>
									</td>
								</tr>
								<tr>
									<td class='header' colspan='10'><nobr>Contact Information </td>
								</tr>
								<tr>
									<td align='right'>Request ID:</td>
									<td class='confirm' width='300px'><%=Request("ID")%>&nbsp;<%=tmpEmer%></td>
								</tr>
								<tr>
									<td align='right'>Timestamp:</td>
									<td class='confirm' width='300px'><%=TS%></td>
								</tr>
								<tr>
									<td align='right'>Status:</td>
									<td class='confirm' width='300px'><%=Statko%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Institution:</td>
									<td class='confirm'><%=tmpIname%></td>
								</tr>
								<tr>
									<td align='right'>Department:</td>
									<td class='confirm'><%=tmpDname%></td>
								</tr>
								<tr>
									<td align='right'>Address:</td>
									<td class='confirm'><%=tmpDeptaddr%></td>
								</tr>
								<tr>
									<td align='right'>Billed To:</td>
									<td class='confirm'><%=tmpBContact%></td>
								</tr>
								<tr>
									<td align='right'>Billing Address:</td>
									<td class='confirm'><%=tmpBaddr%></td>
								</tr>
								<% If Request.Cookies("LBUSERTYPE") <> 4 Then %>
									<tr>
										<td align='right' width='15%'>Rate:</td>
										<td class='confirm'><%=tmpInstRate%></td>
									</tr>
								<% End If %>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Requesting Person:</td>
									<td class='confirm'><%=tmpRP%></td>
								</tr>
								<tr>
									<td align='right'>Phone:</td>
									<td class='confirm'><%=fon%></td>
								</tr>
								<tr>
									<td align='right'>Fax:</td>
									<td class='confirm'><%=fax%></td>
								</tr>
								<tr>
									<td align='right'>E-Mail:</td>
									<td class='confirm'><%=email%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Appointment Information</td>
								</tr>
								<tr>
										<td align='right'>*Client Last Name:</td>
										<td>
											<input class='main' size='20' maxlength='20' name='txtClilname' value='<%=tmplname%>' onkeyup='bawal(this);'>&nbsp;First Name:
											<input class='main' size='20' maxlength='20' name='txtClifname' value='<%=tmpfname%>' onkeyup='bawal(this);'>
										</td>
										<td align='right'>LSS Client:</td>
											<td><input type='checkbox' name='chkClient' value='1' <%=chkClient%>></td>
									</tr>
									<tr>
										<td align='right'>Apartment/Suite Number:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtCliAddrI' value='<%=tmpCAdrI%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right'><nobr>Alternate Appointment Address:</td>
										<td colspan='3'><nobr>
											<input class='main' size='50' maxlength='50' name='txtCliAdd' value='<%=tmpAddr%>' onkeyup='bawal(this);'>
											<input type='checkbox' name='chkClientAdd' value='1' <%=chkUClientadd%>>Check this box and Use these fields if appointment is different from above
											<br>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Do not include apartment, floor, suite, etc. numbers</span>
										</td>
									</tr>
									<tr>
										<td align='right'>City:</td>
										<td>
											<input class='main' size='25' maxlength='25' name='txtCliCity' value='<%=tmpCity%>' onkeyup='bawal(this);'>&nbsp;State:
											<input class='main' size='2' maxlength='2' name='txtCliState' value='<%=tmpState%>' onkeyup='bawal(this);'>&nbsp;Zip:
											<input class='main' size='10' maxlength='10' name='txtCliZip' value='<%=tmpZip%>' onkeyup='bawal(this);'>
										</td>
										<td align='right'>Client Phone:</td>
										<td><input class='main' size='12' maxlength='12' name='txtCliFon' value='<%=tmpCFon%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right'>Gender:</td>
										<td>
											<select class='seltxt' name='selGender' style='width: 75px;'>
												<option value='0' <%=tmpMale%>>Male</option>
												<option value='1' <%=tmpfeMale%>>Female</option>
											</select>
											&nbsp;&nbsp;
											Minor:
											<input type='checkbox' name='chkMinor' value='1' <%=chkMinor%>>
										</td>
										<td align='right'>Alter. Phone:</td>
										<td align='left' rowspan='2'>
											<textarea name='txtAlter' class='main' onkeyup='bawal(this);' ><%=tmpCAFon%></textarea>
										</td>
									</tr>
									<tr>
										<td align='right'>Directions / Landmarks:</td>
										<td><input class='main' size='50' maxlength='50' name='txtCliDir' value='<%=tmpDir%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right'>Special Circumstances:</td>
										<td><input class='main' size='50' maxlength='50' name='txtCliCir' value='<%=tmpSC%>' onkeyup='bawal(this);'></td>
										<td>&nbsp;</td>
										<%If tmpHPID <> 0 Then%>
											<td><u>HospitalPilot Information:</u></td>
										<%End If%>
									</tr>
									<tr>
										<td align='right'>DOB:</td>
										<td>
											<input class='main' size='11' maxlength='10' name='txtDOB' value='<%=tmpDOB%>' onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');">
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
										</td>
										<%If tmpHPID <> 0 Then%>
											<td>&nbsp;</td>
											<td rowspan='4' valign='top'>
												<table cellSpacing='2' cellPadding='0'  border='0' style='border:2px solid;'>
													<tr>
														<td align='right'>ID:</td>
														<td>
															<input class='main' size='11' maxlength='10' name='txtHPID' readonly value='<%=Z_ZeroToNull(tmpHPID)%>' onkeyup='bawal(this);'>
															<input type='hidden' name='hideHPID' value='<%=Z_CZero(tmpHPID)%>'>
														</td>
													</tr>
													<tr>
														<td align='right' valign='top'>Requester's Name:</td>
														<td class='confirm'><%=tmpReqname%></td>
													</tr>
													<tr>
														<td align='right' valign='top'><nobr>Reason:</td>
														<td class='confirm'><%=tmpReason%></td>
													</tr>
													<tr>
														<td align='right'><nobr>Clinician:</td>
														<td class='confirm'><%=tmpClin%></td>
													</tr>
													<%If InHP = 1 Then%>
														<tr>
															<td align='right'><nobr>Meeting Place:</td>
															<td class='confirm'><%=tmpMeet%></td>
														</tr>
													<%End If%>
													<%If tmpParents <> "" Then%>
														<tr>
															<td align='right'><nobr>Parents:</td>
															<td class='confirm'><%=tmpParents%></td>
														</tr>
													<%End If%>
													<tr>
														<td class='confirm' colspan='2'><%=tmpMinor%></td>
													</tr>
													<tr>
														<td class='confirm' colspan='2'><%=tmpCallMe%></td>
													</tr>
												</table>
											</td>
										<%End If%>
									</tr>
									<tr>
										<td align='right'>*Language:</td>
										<td>
											<select class='seltxt' name='selLang'  style='width:100px;' onchange=''>
												<option value='-1'>&nbsp;</option>
												<%=strLang%>
											</select>
											<input type='hidden' name='myLang' value='<%=tmpLang%>'>
										</td>
									</tr>
									<tr>
										<td align='right'>*Appointment Date:</td>
										<td>
											<input class='main' size='10' maxlength='10' name='txtAppDate'  readonly value='<%=tmpAppDate%>'>
											<input type="button" value="..." title='Calendar' name="cal1" style="width: 19px;"
											onclick="showCalendarControl(document.frmAssign.txtAppDate);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
											<input type='hidden' name='mydate' value='<%=tmpAppDate%>'>
										</td>
									</tr>
									<tr>
										<td align='right'>*Appointment Time:</td>
										<td>
											&nbsp;From:<input class='main' size='5' maxlength='5' name='txtAppTFrom' value='<%=tmpAppTFrom%>' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');">
											&nbsp;To:<input class='main' size='5' maxlength='5' name='txtAppTTo' value='<%=tmpAppTTo%>' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');">
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">24-hour format</span>
											<input type='hidden' name='mystime' value='<%=tmpAppTFrom%>'>
										</td>
									</tr>
									<tr>
										<td align='right'>Appointment Location:</td>
										<td><input class='main' size='50' maxlength='50' name='txtAppLoc' value='<%=tmpAppLoc%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right'><b>For legal appointments:</b></td>
										<td><b>(also fill in)</b></td>
									</tr>
									<tr>
										<td align='right'>Docket Number:</td>
										<td><input class='main' size='50' maxlength='50' name='txtDocNum' value='<%=tmpDoc%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right'>Court Room No:</td>
										<td><input class='main' size='12' maxlength='12' name='txtCrtNum' value='<%=tmpCRN%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>	
										<td align='right' valign='top'>Appointment Comment:</td>
										<td colspan='3' >
											<textarea name='txtcom' class='main' onkeyup='bawal(this);' style='width: 375px;'><%=tmpCom%></textarea>
										</td>
									</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Interpreter Information</td>
								</tr>
								<tr>
									<td align='right'>Interpreter:</td>
									<td class='confirm'><%=tmpIntrName%>
										<input type='hidden' name='myint' value='<%=tmpIntr%>'>
										</td>
								</tr>
								<tr>
									<td align='right'>Rate:</td>
									<td class='confirm'><%=tmpIntrRate%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Interpreter Comment:</td>
									<td class='confirm'><%=tmpComintr%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
										<tr><td>&nbsp;</td></tr>
									<tr>
									<td colspan='10' class='header'><nobr>Billing Information</td>
								</tr>
								<tr>
									<td align='right'>Billable Hours:</td>
									<td class='confirm'><%=tmpBilHrs%></td>
								</tr>
								<tr>
									<td align='right'>Actual Time:</td>
									<td class='confirm'><%=tmpActTFrom%> - <%=tmpActTTo%></td>
								</tr>
								<tr>
									<td align='right'>&nbsp;</td>
									<td rowspan='3' valign='top'>
										<table cellSpacing='2' cellPadding='0' border='0'>
											<tr>
												<td align='left'>Bill To Institution </td>
												<td>|</td>
												<td>Pay To Interpreter</td>
											</tr>
											<tr>
												<td class='confirm' align='center'><%=tmpBilTInst%></td>
												<td>|</td>
												<td class='confirm' align='center'><%=tmpBilTIntr%></td>
											</tr>
											<tr>
												<td class='confirm' align='center'><%=tmpBilMInst%> </td>
												<td>|</td>
												<td class='confirm' align='center'> <%=tmpBilMIntr%></td>
											</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td align='right'>Travel Time:</td>
								</tr>
								<tr>
									<td align='right'>Mileage:</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Billing Comment:</td>
									<td class='confirm'><%=tmpCombil%></td>
								</tr>		
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Language Bank Notes	
									</td>
								</tr>
								<tr>
									<td align='right' valign='top'>Notes:</td>
									<td class='confirm'>
										<textarea name='txtLBcom' class='main' onkeyup='bawal(this);' style='width: 375px;' readonly><%=tmpLBcom%></textarea>
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
										<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' align='center' height='100px' valign='bottom'>
												<input type='hidden' name="HID" value='<%=Request("ID")%>'>
												<input type='hidden' name="hidInstRate" value='<%=tmpInstRate%>'>
												<input type='hidden' name="hidIntrRate" value='<%=tmpIntrRate%>'>
												<input class='btn' type='button' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveAss(<%=Request("ID")%>) ;'>
												<input class='btn' type='button' value='Back' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='reqconfirm.asp?ID=<%=Request("ID")%>';">
												<input class='btn' type='button' value='Delete' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='KillMe(<%=Request("ID")%>);'>
											</td>
									</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td valign='bottom'>
							<!-- #include file="_footer.asp" -->
						</td>
					</tr>
				</table>
			</form>
		</body>
	</head>
</html>
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>