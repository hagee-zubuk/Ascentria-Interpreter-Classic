<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
	If Session("UIntr") = "" Then 
		Session("MSG") = "ERROR: Session has expired.<br> Please sign in again."
		Response.Redirect "default.asp"
	End If
	If Request.Cookies("ONCALL") <> 1 Then 
		Session("MSG") = "ERROR: You are not allowed on this page.<br> Please sign in again."
		Response.Redirect "default.asp"
	End If
	Function checkOC(IntrID, InstID, strDate)
		checkOC = ""
		If InstID = "" Or InstID = 0 Then Exit Function
		Set rsOC = Server.CreateObject("ADODB.RecordSet")
		sqlOC = "SELECT *  FROM oncall_T WHERE IntrID = " & IntrID & " AND InstID = " & InstID & " AND OCdate = '" & strDate & "' AND PM = 0"
		rsOC.Open sqlOC, g_strCONN, 3, 1
		If Not rsOC.EOF Then tmpcheckOC = "Checked"
		rsOC.Close
		Set rsOC = Nothing
		Set rsOC = Server.CreateObject("ADODB.RecordSet")
		sqlOC = "SELECT *  FROM oncall_T WHERE IntrID <> " & IntrID & " AND InstID = " & InstID & " AND OCdate = '" & strDate & "' AND PM = 0"
		rsOC.Open sqlOC, g_strCONN, 3, 1
		If Not rsOC.EOF Then 
			tmpcheckOC = tmpcheckOC & " Disabled"
		End If
		rsOC.Close
		Set rsOC = Nothing
		checkOC = tmpcheckOC
	End Function
	Function checkOC2(IntrID, InstID, strDate)
		checkOC2 = ""
		If InstID = "" Or InstID = 0 Then Exit Function
		Set rsOC = Server.CreateObject("ADODB.RecordSet")
		sqlOC = "SELECT *  FROM oncall_T WHERE IntrID = " & IntrID & " AND InstID = " & InstID & " AND OCdate = '" & strDate & "' AND PM = 1"
		rsOC.Open sqlOC, g_strCONN, 3, 1
		If Not rsOC.EOF Then tmpcheckOC = "Checked"
		rsOC.Close
		Set rsOC = Nothing
		Set rsOC = Server.CreateObject("ADODB.RecordSet")
		sqlOC = "SELECT *  FROM oncall_T WHERE IntrID <> " & IntrID & " AND InstID = " & InstID & " AND OCdate = '" & strDate & "' AND PM = 1"
		rsOC.Open sqlOC, g_strCONN, 3, 1
		If Not rsOC.EOF Then 
			tmpcheckOC = tmpcheckOC & " Disabled"
		End If
		rsOC.Close
		Set rsOC = Nothing
		checkOC2 = tmpcheckOC
	End Function
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT [index] as myInst, Facility FROM Institution_T WHERE oncall = 1 ORDER BY Facility"
	rsInst.Open sqlInst, g_strCONN, 3, 1
	Do Until rsInst.EOF
		thisInst = ""
		If request("InstID") <> "" Then
			If Z_CZero(request("InstID")) = rsInst("myInst") Then thisInst = "SELECTED"
		End If 
		strInst = strInst & "<option value='" & rsInst("myInst") & "' " & thisInst & ">" & rsInst("Facility") & "</option>" & vbCrLf
		rsInst.MoveNext
	Loop
	rsInst.Close
	Set rsInst = Nothing
	If request("InstID") <> "" Then
		Set rsOC = Server.CreateObject("ADODB.RecordSet")
		sqlOC = "SELECT * FROM oncall_T WHERE IntrID = " & Session("UIntr") & " AND InstID = " & request("InstID") & " ORDER BY OCDate"
		rsOC.Open sqlOC, g_strCONN, 3, 1
		Do Until rsOC.EOF
			'strOC = "<tr><td>" & rsOC("OCDate") & "</td><td>" & rsOC("OCtime") & "</td></tr>" & vbCrLf
			rsOC.MoveNext
		Loop
		rsOC.Close
		Set rsOC = Nothing
	End If
	'SET OC AVAIL
	tmpReqMonth = Request("selMonth")
tmpReqYear = Request("txtyear")
	If tmpReqMonth <> "" And tmpReqYear <> "" Then
		tmp1Day = tmpReqMonth & "/01/" & tmpReqYear
		tmpMonth = MonthName(tmpReqMonth) & " - " & tmpReqYear
		myMonth = tmpReqMonth
		myYear = tmpReqYear
	End If
	If tmp1Day = "" Then 
		tmp1Day = Month(Date) & "/01/" & Year(Date)
		tmpMonth = MonthName(Month(Date)) & " - " & Year(Date)
		myMonth = Month(Date)
		myYear = Year(Date)
	End If
	If Not IsDate(tmp1Day) Then 
		tmp1day = Month(Date) & "/01/" & Year(Date)
		tmpMonth = MonthName(Month(Date)) & " - " & Year(Date)
		myMonth = Month(Date)
		myYear = Year(Date)
		Session("MSG") = "ERROR: Year inputted is not valid. Set to current month and year."
	End If
	CorrectMonth = True
	tmpToday = tmp1Day
	tmp1day2 = dateadd("m", 2, Month(Date) & "/01/" & Year(Date))
	tmp1day22 = Month(Date) & "/01/" & Year(Date)
	
	monthdis = ""
	If cdate(tmptoday) >= cdate(tmp1day2) Then
		monthdis = "disabled"
	End If
	monthdis2 = ""
	If cdate(tmptoday) <= cdate(tmp1day22) Then
		monthdis2 = "disabled"
	End If
	lastday = Day(DateSerial(Year(tmptoday), Month(tmptoday) + 1, 0))
	Do While CorrectMonth = True 
		'set calendar
		strCal = strCAL & "<tr><td>&nbsp;</td>"
		If WeekdayName(Weekday(tmpToday), True) <> "Sun" Then 
			strCal = strCAL & "<td colspan='3'>&nbsp;</td>"
		Else
			strMonth = Month(tmpToday) 
			strDay = Day(tmpToday)
			strYear = Year(tmpToday)
			tmpBG = "#FFFFFF"
			OCchk = checkOC(Session("UIntr"), request("InstID"), tmpToday)
			OCchk2 = checkOC2(Session("UIntr"), request("InstID"), tmpToday)
			If len(Day(tmpToday)) = 1 Then
				strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top'>" & Day(tmpToday) & "</td><td align='center' class='caltbl'><input type='checkbox' name='chk" & Day(tmpToday) & "' value='1' " & OCchk & "></td><td align='center' class='caltbl'><input type='checkbox' name='chkp" & Day(tmpToday) & "' value='1' " & OCchk2 & "></td>" & vbCrLf
			Else
				strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top'>" & Day(tmpToday) & "</td><td align='center' class='caltbl'><input type='checkbox' name='chk" & Day(tmpToday) & "' value='1' " & OCchk & "></td><td align='center' class='caltbl'><input type='checkbox' name='chkp" & Day(tmpToday) & "' value='1' " & OCchk2 & "></td>" & vbCrLf
			End If
			tmpToday = DateAdd("d", 1, tmpToday)
			If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
		End If	
		If WeekdayName(Weekday(tmpToday), True) <> "Mon" Then 
			strCal = strCAL & "<td>&nbsp;</td>"
		Else
			strMonth = Month(tmpToday) 
			strDay = Day(tmpToday)
			strYear = Year(tmpToday)
			tmpBG = "#FFFFFF"
			OCchk = checkOC(Session("UIntr"), request("InstID"), tmpToday)
			If len(Day(tmpToday)) = 1 Then
				strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top'>" & Day(tmpToday) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='checkbox' name='chk" & Day(tmpToday) & "' value='1' " & OCchk & "></td>" & vbCrLf
			Else
				strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top'>" & Day(tmpToday) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='checkbox' name='chk" & Day(tmpToday) & "' value='1' " & OCchk & "></td>" & vbCrLf
			End If
			tmpToday = DateAdd("d", 1, tmpToday)
			If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
		End If	
		If WeekdayName(Weekday(tmpToday), True) <> "Tue" Then 
			strCal = strCAL & "<td>&nbsp;</td>"
		Else
			strMonth = Month(tmpToday) 
			strDay = Day(tmpToday)
			strYear = Year(tmpToday)
			tmpBG = "#FFFFFF"
			OCchk = checkOC(Session("UIntr"), request("InstID"), tmpToday)
			If len(Day(tmpToday)) = 1 Then
				strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top'>" & Day(tmpToday) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='checkbox' name='chk" & Day(tmpToday) & "' value='1' " & OCchk & "></td>" & vbCrLf
			Else
				strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top'>" & Day(tmpToday) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='checkbox' name='chk" & Day(tmpToday) & "' value='1' " & OCchk & "></td>" & vbCrLf
			End If
			tmpToday = DateAdd("d", 1, tmpToday)
			If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
		End If	
		If WeekdayName(Weekday(tmpToday), True) <> "Wed" Then 
			strCal = strCAL & "<td>&nbsp;</td>"
		Else
			strMonth = Month(tmpToday) 
			strDay = Day(tmpToday)
			strYear = Year(tmpToday)
			tmpBG = "#FFFFFF"
			OCchk = checkOC(Session("UIntr"), request("InstID"), tmpToday)
			If len(Day(tmpToday)) = 1 Then
				strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top'>" & Day(tmpToday) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='checkbox' name='chk" & Day(tmpToday) & "' value='1' " & OCchk & "></td>" & vbCrLf
			Else
				strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top'>" & Day(tmpToday) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='checkbox' name='chk" & Day(tmpToday) & "' value='1' " & OCchk & "></td>" & vbCrLf
			End If
			tmpToday = DateAdd("d", 1, tmpToday)
			If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
		End If	
		If WeekdayName(Weekday(tmpToday), True) <> "Thu" Then 
			strCal = strCAL & "<td>&nbsp;</td>"
		Else
			strMonth = Month(tmpToday) 
			strDay = Day(tmpToday)
			strYear = Year(tmpToday)
			tmpBG = "#FFFFFF"
			OCchk = checkOC(Session("UIntr"), request("InstID"), tmpToday)
			If len(Day(tmpToday)) = 1 Then
				strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top'>" & Day(tmpToday) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='checkbox' name='chk" & Day(tmpToday) & "' value='1' " & OCchk & "></td>" & vbCrLf
			Else
				strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top'>" & Day(tmpToday) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='checkbox' name='chk" & Day(tmpToday) & "' value='1' " & OCchk & "></td>" & vbCrLf
			End If
			tmpToday = DateAdd("d", 1, tmpToday)
			If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
		End If	
		If WeekdayName(Weekday(tmpToday), True) <> "Fri" Then 
			strCal = strCAL & "<td>&nbsp;</td>"
		Else
			strMonth = Month(tmpToday) 
			strDay = Day(tmpToday)
			strYear = Year(tmpToday)
			tmpBG = "#FFFFFF"
			OCchk = checkOC(Session("UIntr"), request("InstID"), tmpToday)
			If len(Day(tmpToday)) = 1 Then
				strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top'>" & Day(tmpToday) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='checkbox' name='chk" & Day(tmpToday) & "' value='1' " & OCchk & "></td>" & vbCrLf
			Else
				strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top'>" & Day(tmpToday) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='checkbox' name='chk" & Day(tmpToday) & "' value='1' " & OCchk & "></td>" & vbCrLf
			End If
			tmpToday = DateAdd("d", 1, tmpToday)
			If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
		End If	
		If WeekdayName(Weekday(tmpToday), True) <> "Sat" Then 
			strCal = strCAL & "<td colspan='3'>&nbsp;</td>"
		Else
			strMonth = Month(tmpToday) 
			strDay = Day(tmpToday)
			strYear = Year(tmpToday)
			tmpBG = "#FFFFFF"
			OCchk = checkOC(Session("UIntr"), request("InstID"), tmpToday)
			OCchk2 = checkOC2(Session("UIntr"), request("InstID"), tmpToday)
			If len(Day(tmpToday)) = 1 Then
				strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top'>" & Day(tmpToday) & "</td><td align='center' class='caltbl'><input type='checkbox' name='chk" & Day(tmpToday) & "' value='1' " & OCchk & "></td><td align='center' class='caltbl'><input type='checkbox' name='chkp" & Day(tmpToday) & "' value='1' " & OCchk2 & "></td>" & vbCrLf
			Else
				strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top'>" & Day(tmpToday) & "</td><td align='center' class='caltbl'><input type='checkbox' name='chk" & Day(tmpToday) & "' value='1' " & OCchk & "></td><td align='center' class='caltbl'><input type='checkbox' name='chkp" & Day(tmpToday) & "' value='1' " & OCchk2 & "></td>" & vbCrLf
			End If
			tmpToday = DateAdd("d", 1, tmpToday)
			If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
		End If	
		strCal = strCAL & "</tr>"
		If Month(tmp1Day) <> Month(tmpToday) Then CorrectMonth = False
	Loop
	billedna = ""
	If request("InstID") = "" Or request("InstID") = 0 Then billedna = "disabled"
%>
<html>
	<head>
		<title>Language Bank - Interpreter On Call Services</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function getDept(xxx)
		{
			document.frmTS.action = "oncall.asp?InstID=" + xxx;
			document.frmTS.submit();
		}	
		function SaveTS()
		{
			document.frmTS.action = "action.asp?ctrl=19";
			document.frmTS.submit();
	
		}
		function ChangeMonth(xxx)
		{
			document.frmTS.action = "action.asp?page=1&ctrl=4&dir=" + xxx;
			document.frmTS.submit();
		}
		//-->
		</script>
	</head>
	<body>
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
								<td class='title' colspan='10' align='center'><nobr> Interpreter On Call Services</td>
								</tr>
								<tr>
									<td  align='center' colspan='12'>
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
									<td align='right' width='150px'>Name:</td>
									<td class='confirm'><%=GetIntr(Session("UIntr"))%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td>&nbsp;</td>
									<td>
										<form name='frmTS' method='POST' action='oncall.asp'>
											<input type='hidden' name='Hmonth' value='<%=tmpMonth%>'>
											<input type='hidden' name='mymonth' value='<%=mymonth%>'>
											<input type='hidden' name='myyear' value='<%=myyear%>'>
											<input type='hidden' name='lastday' value='<%=lastday%>'>
											<input type='hidden' name='qstr' value='<%=Request.ServerVariables("QUERY_STRING")%>'>
										<table>
											<tr>
												<td>
												<select name='selInst' class='seltxt' onchange="getDept(this.value);" onblur="getDept(this.value);">
													<option value='0'>&nbsp;</option>
													<%=strInst%>
												</select>
											
												</td>
											</tr>
											<tr><td>&nbsp;</td></tr>
											<Tr>
												<td>
													<table cellSpacing='0' cellPadding='0' align='center' style='width: 80%;' border='0'>
													<tr>
														<td align='left'>
															<input class='btn' type='button' value='&lt&lt' title='Previous Month' style='width: 25px;' <%=Monthdis2%> onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='ChangeMonth(0);'>
														</td>
														<td colspan='20' align='center' class='calheader'>
															<%=tmpMonth%>
														</td>
														<td align='right'>
															<input class='btn' type='button' value='&gt&gt' title='Next Month' style='width: 25px;' <%=Monthdis%> onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='ChangeMonth(1);'>
														</td>
													</tr>
													<tr>
														<td>&nbsp;</td>
														<td class='calweekday' colspan='3'>Sun</td>
														<td class='calweekday'>Mon</td>
														<td class='calweekday'>Tue</td>
														<td class='calweekday'>Wed</td>
														<td class='calweekday'>Thu</td>
														<td class='calweekday'>Fri</td>
														<td class='calweekday' colspan='3'>Sat</td>
														<td>&nbsp;</td>
													</tr>
													<%=strCal%>
													<tr><td>&nbsp;</td></tr>
													<tr>
														<td colspan='9' align='center'>
															
															
														</td>
													</tr>
												</table>
												</td>
											</tr>	
										</table>	
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='12' align='center'>
										*To <b>SAVE</b> your desired oncall date click on the corresponding checkbox then click on the 'Save' button.<br>
										*To <b>REMOVE</b> your desired oncall date click on the corresponding checkbox to uncheck then click on the 'Save' button.<br>
										<br><br>
										*Sundays and Saturdays have 2 shifts (am/pm).<br>
										*Disabled checkbox = Already asigned to someone else.
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='12' align='center' height='100px' valign='bottom'>
										<input class='btn' type='button' value='Save' <%=billedna%> onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveTS();'>
									</td>
								</tr>
								</form>
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
<!-- #include file="_closeSQL.asp" -->
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