<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%Response.AddHeader "Pragma", "No-Cache" %>
<%
server.scripttimeout = 360000
'USER CHECK
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
Function GetPrime(xxx)
	GetPrime = ""
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT * FROM requester_T WHERE [index] = " & xxx
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		If rsRP("prime") = 0 Then
			GetPrime = rsRP("Email")
		ElseIf rsRP("prime") = 1 Then
			'GetPrime = rsRP("Phone")
			GetPrime = ""
		ElseIf rsRP("prime") = 2 Then
			GetPrime = CleanFax(Trim(rsRP("Fax"))) & "@emailfaxservice.com" 
		End If
	End If
	rsRP.Close
	set rsRP = Nothing
End Function
tmpPage = "document.frmMain."
tmpInst = "-1"
tmpIntr = "-1"
'default
selRPEmail = ""
selRPPhone = ""
selRPFax = "checked"
'default
selIntrFax = "checked"
selIntrP2 = ""
selIntrP1 = ""
selIntrEmail = ""
tmpTS = Now
tmpDept = 0
tmpReqP = 0
tmpHPID = 0
If Session("MSG") <> "" Then
on error resume next
	tmpEntry = Split(Z_DoDecrypt(Request.Cookies("LBREQUESTW4")), "|")
	tmplName = tmpEntry(1)
	tmpfName = tmpEntry(2)
	chkClient = ""
	If tmpEntry(3) <> "" Then chkClient = "checked"
	tmpCAdrI = tmpEntry(4)
	tmpAddr = tmpEntry(5)
	chkUClientadd = ""
	If tmpEntry(6) <> "" Then chkUClientadd = "checked"
	tmpCFon = tmpEntry(7)
	tmpCity = tmpEntry(8)
	tmpState = tmpEntry(9)
	tmpZip = tmpEntry(10)
	tmpCAFon = tmpEntry(11)
	tmpDir = tmpEntry(12)
	tmpSC = tmpEntry(13)
	tmpDOB = Z_FormatTime(tmpEntry(14))
	tmpLang = tmpEntry(15)
	tmpAppDate = Z_FormatTime(tmpEntry(16))
	tmpAppTFrom = Z_FormatTime(tmpEntry(17))
	tmpAppTTo = Z_FormatTime(tmpEntry(18))
	tmpAppLoc = tmpEntry(19)
	tmpDoc = tmpEntry(20)
	tmpCRN = tmpEntry(21)
	tmpCom = tmpEntry(22)
	'tmpGender = tmpEntry(23)
	'tmpMinor = tmpEntry(24)
	tmpGender	= Z_CZero(tmpEntry(23))
	tmpMale = ""
	tmpFemale = ""
	If tmpGender = 0 Then 
		tmpMale = "SELECTED"
	Else
		tmpFemale = "SELECTED"
	End If
	chkMinor = ""
	If tmpEntry(24) <> "" Then chkMinor = "CHECKED"
End If
If Request("tmpID") <> "" Then
	Set rsW1 = Server.CreateObject("ADODB.RecordSet")
	sqlW1 = "SELECT * FROM Wrequest_T WHERE [index] = " & Request("tmpID")
	rsW1.Open sqlW1, g_strCONNW, 1, 3
	If Not rsw1.EOF Then
		myInst = rsW1("InstID")
		tmpEmer = ""
		If rsW1("Emergency") = True Then tmpEmer = "checked"
		tmpEmerFee = ""
		If rsW1("EmerFee") = True Then tmpEmerFee = "checked"
		tmpInstRate = rsW1("InstRate")
		tmpDept = rsW1("DeptID")
		tmpReqP = rsW1("ReqID")
	End If
	rsW1.Close
	Set rsW1 = Nothing
End If
'GET TEMP DATA
Set rsWdata = Server.CreateObject("ADODB.RecordSet")
sqlWdata = "SELECT * FROM Wrequest_T WHERE [index] = " & Request("tmpID")
rsWdata.Open sqlWdata, g_strCONNW, 1, 3
If Not rsWdata.EOF Then
	tmpInst = rsWdata("instID")
	tmpEmer = ""
	If rsWdata("Emergency") = True Then tmpEmer = "(EMERGENCY)" 
	tmpInstRate = Z_FormatNumber(rsWdata("InstRate"), 2)	
End If
rsWdata.Close
Set rsWdata = Nothing
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
		'tmpDeptaddr = rsDept("InstAdrI") & " " & rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
		tmpDeptaddrG = rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	End If
End If
rsDept.Close
Set rsDept = Nothing 
'GET REQUESTING PERSON
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM requester_T WHERE [index] = " & tmpReqP
rsReq.Open sqlReq, g_strCONN, 3, 1
If Not rsReq.EOF Then
	tmpRP = rsReq("Lname") & ", " & rsReq("Fname") 
	Fon = rsReq("phone") 
	If rsReq("pExt") <> "" Then Fon = Fon & " ext. " & rsReq("pExt")
	Fax = rsReq("fax")
	email = rsReq("email")
	Pcon = GetPrime(tmpReqP)
End If
rsReq.Close
Set rsReq = Nothing
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
		<title>Language Bank - Interpreter Request Form - Appointment Information</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
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
		function CalendarView(strDate)
		{
			document.frmMain.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmMain.submit();
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
		function WSubmit(xxx)
		{
			if (document.frmMain.chkClientAdd.checked == true)
			{
				if (document.frmMain.txtCliAdd.value == "" || document.frmMain.txtCliCity.value == "" || document.frmMain.txtCliState.value == "" || document.frmMain.txtCliZip.value == "")
				{
					alert("Please input client's full address.")
					return;
				}
			}
			if (document.frmMain.txtClilname.value == "" && document.frmMain.txtClifname.value == "")
			{
				alert("ERROR: Client is Required."); 
				return;
			}
			if (document.frmMain.selLang.value == 0)
			{
				alert("ERROR: Language is Required."); 
				return;
			}
			if (document.frmMain.txtAppDate.value == "")
			{
				alert("ERROR: Appointment Date is Required."); 
				return;
			}
			if (document.frmMain.txtAppTFrom.value == "")
			{
				alert("ERROR: Appointment Time (From:) is Required."); 
				return;
			}
			var ans = window.confirm("Submit Appointment to Database?");
			if (ans){
				document.frmMain.action = "waction.asp?ctrl=4";
				document.frmMain.submit();
			}
		}
		function WBack(xxx)
		{
			var ans = window.confirm("Any changes made in this page will not be saved.");
			if (ans){
				document.frmMain.action = "wMain3.asp?tmpID=" + xxx;
				document.frmMain.submit();
			}
		}
		//-->
		</script>
		</head>
		<body onload=''>
			<form method='post' name='frmMain' action='main.asp'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td height='100px'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top' >
							<form name='frmService' method='post' action=''>
								<table cellSpacing='2' cellPadding='0' width="100%" border='0'>
									<!-- #include file="_greetme.asp" -->
									<tr>
										<td class='title' colspan='10' align='center'><nobr> Interpreter Request Form - 4 / 4</td>
									</tr>
									<tr>
										<td align='center' colspan='10'><nobr>(*) required</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td  align='left'>
											<div name="dErr" style="width:100%; height:55px;OVERFLOW: auto;">
												<table border='0' cellspacing='1'>		
													<tr>
														<td><span class='error'><%=Session("MSG")%></span></td>
													</tr>
												</table>
											</div>
										</td>
									</tr>
									<tr>
										<td class='header' colspan='10'><nobr>Contact Information</td>
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
									<tr><td>&nbsp;</td></tr>
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
									
									</tr>
									<tr>
										<td align='right'>DOB:</td>
										<td>
											<input class='main' size='11' maxlength='10' name='txtDOB' value='<%=tmpDOB%>' onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');">
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
										</td>
										
									</tr>
									<tr>
										<td align='right'>*Language:</td>
										<td>
											<select class='seltxt' name='selLang'  style='width:100px;' onchange=''>
												<option value='0'>&nbsp;</option>
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
											onclick="showCalendarControl(document.frmMain.txtAppDate);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
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
									<tr>
										<td colspan='10' align='center' height='100px' valign='bottom'>
											<input class='btn' type='button' value='<<' style='width: 50px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='WBack(<%=Request("tmpID")%>);'>
											<input class='btn' type='Reset' value='Clear' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
											<input class='btn' type='button' value='Cancel' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="window.location='calendarview2.asp'">
											<input class='btn' type='button' value='Submit' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='WSubmit(<%=Request("tmpID")%>);'>
											<input type='hidden' name='tmpID' value='<%=Request("tmpID")%>'>
											<input type='hidden' name='tmpInst' value='<%=tmpInst%>'>
											<input type='hidden' name='tmpDep' value='<%=tmpDept%>'>
											<input type='hidden' name='tmpReqP' value='<%=tmpReqP%>'>
										</td>
									</tr>
									
								</table>
							</form>
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
