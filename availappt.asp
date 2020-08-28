<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
frmDte = Date
toDte = DateAdd("m", 2, Date)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	If Z_DateNull(Request("frmdte")) >= Date And Z_DateNull(Request("frmdte")) <= DateAdd("m", 2, Date) Then frmDte = Request("frmdte")
	If Z_DateNull(Request("toDte")) >= Date And Z_DateNull(Request("toDte")) <= DateAdd("m", 2, Date) Then toDte = Request("toDte")
End If
'get open appt
Set rsApp = Server.CreateObject("ADODB.RecordSet")
'sqlApp = "SELECT appID, appdate, langID, deptID, appTimeFrom, appTimeTo, CliAdd, ccity, cstate, state, accept FROM appt_T, request_T WHERE appt_T.IntrID = " & _
'	Session("UIntr") & " AND appt_T.[appID] = request_T.[index] AND [status] <> 3 AND NOT request_T.intrID > 0 " & _
'	"AND appDate >= '" & frmDte & "' AND appdate <= '" & toDte & "' ORDER BY [accept], appDate"
sqlApp = "SELECT req.[intrID], app.[IntrID], CliAdd, [ccity], [cstate], req.[is_rmt]" & _
		", dep.[city], dep.[state], appID, appdate, langID, deptID, appTimeFrom, appTimeTo, [accept] " & _
		"FROM [appt_T] AS app " & _
		"INNER JOIN [request_T] AS req ON app.[appID]=req.[index] " & _
		"INNER JOIN [dept_T] AS dep ON req.[deptid]=dep.[index] " & _
		"WHERE [status] <> 3 " & _
		"AND req.[intrID] <= 0 " & _
		"AND appDate >= '" & frmDte & "' " & _
		"AND appdate <= '" & toDte & "' " & _
		"AND app.[IntrID] =" & Session("UIntr") & _
		" ORDER BY [accept], [appDate]"	
rsApp.Open sqlApp, g_strCONN, 3, 1
x = 0
Do Until rsApp.EOF
	kulay = ""
	If Z_IsOdd(x) Then kulay = "#FBEEB7"
	If rsApp("is_rmt") Then kulay = "#a2ff84"
	' &#x1f57f;
	'tmpIname = GetInst(Z_GetInfoFROMAppID(rsApp("appID"), "instID"))
	'myDept = GetDept(Z_GetInfoFROMAppID(rsApp("appID"), "deptID"))
	tmpSalita = GetLang(rsApp("LangID")) '(Z_GetInfoFROMAppID(rsApp("appID"), "langID"))
	appdate = rsApp("appdate")
	myclass = GetClass(Z_GetClass(rsApp("deptID")))
	timeframe = Z_FormatTime(rsApp("apptimeFrom"), 4) & " - " & Z_FormatTime(rsApp("appTimeTo"), 4)
	If rsApp("CliAdd") Then
		tmpcity = rsApp("ccity") & ", " & rsApp("cstate")
	Else
		tmpcity = rsApp("city") & ", " & rsApp("state")
	End If
	
	noans = ""
	ansyes = ""
	ansno = ""
	If rsApp("Accept") = 0 Then noans = "selected"
	If rsApp("Accept") = 1 Then ansyes = "selected"
	If rsApp("Accept") = 2 Then ansno = "selected"
	strtbl = strtbl & "<tr bgcolor='" & kulay & "'>" & vbCrLf & _ 
		"<td class='tblgrn2' ><input type='hidden' name='hid" & x & "' value='" & rsApp("appID") & "' ><b>" & rsApp("appID") & "</b></td>" & vbCrLf & _
		"<td class='tblgrn2' >" & myclass & "</td>" & vbCrLf & _
		"<td class='tblgrn2' >" & tmpCity & "</td>" & vbCrLf & _
		"<td class='tblgrn2' >" & tmpSalita & "</td>" & vbCrLf & _
		"<td class='tblgrn2' >" & appdate & "</td>" & vbCrLf & _
		"<td class='tblgrn2' >" & timeframe & "</td>" & vbCrLf & _
		"<td class='tblgrn2' ><input class='btntbl' type='button' value='Payable Travel Time/Mileage' style='width: 180px;' onmouseover=""this.className='hovbtntbl'"" onmouseout=""this.className='btntbl'"" onclick='CheckGoogle(" & rsApp("appID") & ", " & Session("UIntr") & ");'></td>" & vbCrLf & _
		"<td class='tblgrn2' >"
	'If rsApp("Accept") = 0 Then
		strtbl = strtbl & "<select name='selSagot" & x & "' class='seltxt' style='width:50px;'>" & _
				"<option value='0' " & noans & ">&nbsp;</option>" & _
				"<option value='1' " & ansyes & ">Yes</option>" & _
				"<option value='2' " & ansno & ">No</option>" & _
			"</select>"
	'ElseIf rsApp("Accept") = 1 Then
	'	strtbl = strtbl & "<input type='hidden' name='selSagot" & x & "' value='1' >Yes"
	'ElseIf rsApp("Accept") = 2 Then
	'	strtbl = strtbl & "<input type='hidden' name='selSagot" & x & "' value='2' >No"
	'End If
	strtbl = strtbl & "</td>" & vbCrlf & "</tr>" & vbCrlf
	x = x + 1

	rsApp.MoveNext
Loop
rsApp.Close
Set rsApp = Nothing
<!-- #include file="_closeSQL.asp" -->
%>
<html>
	<head>
		<title>Language Bank - Open Appointments <%=Session("UIntr")%></title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
			function SaveAns() {
				var ans = window.confirm("Being available does not mean you are assigned to that appointment.\nPlease wait for the confirmation that you were assigned.\n\nSave Availability?\nPlease double check your enties.\nClick Cancel to stop.");
				if (ans) {
					document.frmTbl.action = "saveAvail.asp"
					document.frmTbl.submit();
				}
			}
			function CheckGoogle(appID, intrID) {
				newwindow = window.open('FindMileageTravel.asp?appID=' + appID + '&intrID=' + intrID,'name','height=150,width=400,scrollbars=0,directories=0,status=0,toolbar=0,resizable=0');
				if (window.focus) {newwindow.focus()}
			}
			function FindOpen(frmdte, todte) {
			document.frmTbl.action = "availappt.asp?frmdte=" + frmdte + "&todte=" + todte;
			document.frmTbl.submit();
		}
		//-->
		</script>
		<style type="text/css">
	 	.container
	      {
	          border: solid 1px black;
	          overflow: auto;
	      }
	      .noscroll
	      {
	          position: relative;
	          background-color: white;
	          top:expression(this.offsetParent.scrollTop);
	      }
	      th
	      {
	          text-align: left;
	      }
		</style>
		<body >
			<form method='post' name='frmTbl' action='availappt.asp'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td height='100px'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top'>
							<table cellSpacing='0' cellPadding='0' width="100%" border='0'>
									<!-- #include file="_greetme.asp" -->
								<tr>
									<td>
										<table cellpadding='0' cellspacing='0' width='100%' border='0'>
											<tr>
												<td colspan='10' align='left'>
													<table class="reqtble">	
														<tr>
															<td align='right'>Date Range:</td>
															<td>
																<input class='main' size='10' maxlength='10' name='txtFromDate'  readonly value='<%=frmdte%>'>
																<input type="button" value="..." title='Calendar' name="calFrom" style="width: 19px;"
																onclick="showCalendarControl(document.frmTbl.txtFromDate);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
																&nbsp;&nbsp;-&nbsp;&nbsp;
																<input class='main' size='10' maxlength='10' name='txtToDate'  readonly value='<%=todte%>'>
																<input type="button" value="..." title='Calendar' name="calTo" style="width: 19px;"
																onclick="showCalendarControl(document.frmTbl.txtToDate);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
																&nbsp;&nbsp;
																<input type='button' value='Search' name='btnSearch' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='FindOpen(document.frmTbl.txtFromDate.value, document.frmTbl.txtToDate.value);'>
																* Search will display appointments whithin 2 months from today (<%=Date%> - <%=DateAdd("m", 2, Date)%>) ONLY.
															</td>
														</tr>
													</table>
												</td>
												
												<td align='right'>
													<input type='hidden' name='Hctr' value='<%=x%>'>
													<input class='btntbl' type='button' value='Save' style='height: 25px; width: 150px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='SaveAns();'>
												</td>
											</tr>
										</table>
									</td>
								</tr>
								<% If Session("MSG") <> "" Then %>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='14' align='left'>
											<div name="dErr" style="width:300px; height:40px;OVERFLOW: auto;">
												<table border='0' cellspacing='1'>		
													<tr>
														<td><span class='error'><%=Session("MSG")%></span></td>
													</tr>
												</table>
											</div>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
								<% End If %>
								<tr>
									<td colspan='10' align='left'>
										<div class='container' style='height: 500px; width:1000px; position: relative;'>
											<table class="reqtble" width='100%' >	
												<thead>
													<tr class="noscroll">	
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Request ID</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Type</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">City</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Language</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Appointment Date</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Planned Start and End Time</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">&nbsp;</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Available?</td>
													</tr>
												</thead>
												<tbody style="OVERFLOW: auto;">
													<%=strtbl%>
												</tbody>
											</table>
										</div>	
										<p><font color='#a2ff84' size='2'>&#x2588;</font>&nbsp;-&nbsp;Remote</p>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table width='100%'  border='0'>
								<tr>
									<td align='left'>
										* Being available does not mean you are assigned to that appointment. Please wait for the confirmation that you were assigned.
									</td>
									<td align='right'>
										<% If x <> 0 Then %>
											<b><u><%=x%></u></b> record/s &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<% End If %>
									</td>
									<td>&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td height='50px' valign='bottom'>
							<!-- #include file="_footer.asp" -->
						</td>
					</tr>
				</table>
			</form>
			<code>
			<%=sqlApp%>
			</code>
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