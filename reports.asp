<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) = 2 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
tmpPage = "document.frmReport."
tmpRep = Request("rep")
If Request("rep") = "" Then tmpRep = 0
tmpSele = Request("sel")
If Request("sel") = "" Then tmpSele = 0
'ON ERROR
If Session("MSG") <> "" Then
	tmpReport = Split(Z_DoDecrypt(Request.Cookies("LBREPORT")), "|")
	TypeSel1 = ""
	TypeSel2 = ""
	TypeSel3 = ""
	TypeSel4 = ""
	TypeSel5 = ""
	TypeSel6 = ""
	TypeSel7 = ""
	TypeSel8 = ""
	TypeSel9 = ""
	TypeSel10 = ""
	TypeSel11 = ""
	TypeSel12 = ""
	TypeSel13 = ""
	TypeSel14 = ""
	TypeSel15 = ""
	TypeSel16 = ""
	TypeSel17 = ""
	TypeSel18 = ""
	TypeSel19 = ""
	TypeSel20 = ""
	TypeSel21 = ""
	TypeSel22 = ""
	TypeSel23 = ""
	TypeSel24 = ""
	TypeSel25 = ""
	TypeSel26 = ""
	TypeSel27 = ""
	TypeSel28 = ""
	If tmpReport(0) = 1 Then TypeSel1 = "selected"
	If tmpReport(0) = 2 Then TypeSel2 = "selected"
	If tmpReport(0) = 3 Then TypeSel3 = "selected"
	If tmpReport(0) = 4 Then TypeSel4 = "selected"
	If tmpReport(0) = 5 Then TypeSel5 = "selected"
	If tmpReport(0) = 6 Then TypeSel6 = "selected"
	If tmpReport(0) = 7 Then TypeSel7 = "selected"
	If tmpReport(0) = 8 Then TypeSel8 = "selected"
	If tmpReport(0) = 9 Then TypeSel9 = "selected"
	If tmpReport(0) = 10 Then TypeSel10 = "selected"
	If tmpReport(0) = 11 Then TypeSel11 = "selected"
	If tmpReport(0) = 12 Then TypeSel12 = "selected"
	If tmpReport(0) = 13 Then TypeSel13 = "selected"
	If tmpReport(0) = 14 Then TypeSel14 = "selected"
	If tmpReport(0) = 15 Then TypeSel15 = "selected"
	If tmpReport(0) = 16 Then TypeSel16 = "selected"
	If tmpReport(0) = 17 Then TypeSel17 = "selected"
	If tmpReport(0) = 18 Then TypeSel18 = "selected"
	If tmpReport(0) = 19 Then TypeSel19 = "selected"
	If tmpReport(0) = 20 Then TypeSel20 = "selected"
	If tmpReport(0) = 21 Then TypeSel21 = "selected"
	If tmpReport(0) = 22 Then TypeSel22 = "selected"
	If tmpReport(0) = 23 Then TypeSel23 = "selected"
	If tmpReport(0) = 24 Then TypeSel24 = "selected"
	If tmpReport(0) = 25 Then TypeSel25 = "selected"
	If tmpReport(0) = 26 Then TypeSel26 = "selected"
	If tmpReport(0) = 27 Then TypeSel27 = "selected"
	If tmpReport(0) = 28 Then TypeSel28 = "selected"
	tmpRepFrom = tmpReport(1)
	tmpRepTo = tmpReport(2)
	tmpInst = Z_Cdbl(tmpReport(3))
	tmpIntr = tmpReport(4)
	tmpTown = tmpReport(5)
End If
strIntr = ""
strTown = ""
strLang = ""
strCli = ""
'GET INTERPRETER
'Set rsIntr = Server.CreateObject("ADODB.RecordSet")
'sqlIntr = "SELECT * FROM interpreter_T ORDER BY [Last Name], [First Name]"
'rsIntr.Open sqlIntr, g_strCONN, 3, 1
'Do Until rsIntr.EOF
'	tmpSel = ""
'	If tmpIntr = rsIntr("index") Then tmpSel = "selected"
'	strIntr = strIntr & "<option " & tmpSel & " value='" & rsIntr("index") & "'>" & rsIntr("Last Name") & ", " & rsIntr("First Name") & "</option>" & vbCrLf 
'	rsIntr.MoveNext
'Loop
'rsIntr.Close
'Set rsInt = Nothing
'GET TOWNS
'Set rsTown = Server.CreateObject("ADODB.RecordSet")
'sqlTown = "SELECT DISTINCT(city) FROM dept_T ORDER BY city"
'rsTown.Open sqlTown, g_strCONN, 3, 1
'Do Until rsTown.EOF
'	tmpSel = ""
'	If tmpTown = rsTown("City") Then tmpSel = "selected"
'	strTown = strTown & "<option " & tmpSel & " value='" & Trim(rsTown("city")) & "'>" & Trim(rsTown("city")) & "</option>" & vbCrLf 
'	rsTown.MoveNext
'Loop
'rsTown.Close
'Set rsTown = Nothing
'GET INSTITUTION
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T ORDER BY [Facility]"
rsInst.Open sqlInst, g_strCONN, 3, 1
Do Until rsInst.EOF
	tmpSel = ""
	If tmpInst = rsInst("index") Then tmpSel = "selected"
		InstName = rsInst("Facility")
		strInst = strInst	& "<option " & tmpSel & " value='" & rsInst("Index") & "'>" &  InstName & "</option>" & vbCrlf
	rsInst.MoveNext
Loop
rsInst.Close
Set rsInst = Nothing
'GET AVAILABLE LANGUAGES
'Set rsLang = Server.CreateObject("ADODB.RecordSet")
'sqlLang = "SELECT * FROM language_T ORDER BY [Language]"
'rsLang.Open sqlLang, g_strCONN, 3, 1
'Do Until rsLang.EOF
'	tmpLang = Request("LangID")
'	tmpL = ""
'	If tmpLang = "" Then tmpLang = "-1"
'	If CInt(tmpLang) = rsLang("index") Then tmpL = "selected"
'	strLang = strLang	& "<option " & tmpL & " value='" & rsLang("Index") & "'>" &  rsLang("language") & "</option>" & vbCrlf
'	rsLang.MoveNext
'Loop
'rsLang.Close
'Set rsLang = Nothing
'GET CLIENT LIST
'Set rsCli = Server.CreateObject("ADODB.RecordSet")
'sqlCli = "SELECT DISTINCT Clname, Cfname FROM request_T ORDER BY Clname, Cfname"
'rsCli.Open sqlCli, g_strCONN, 3, 1
'Do Until rsCli.EOF
'	strCli = strCli	& "<option>" & rsCli("Clname") & ", " & rsCli("Cfname") & "</option>" & vbCrlf
'	rsCli.MoveNext
'Loop
'rsCli.Close
'Set rsCli = Nothing
todaydate = Cdate(date)
%>
<html>
	<head>
		<title>Language Bank - Reports</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function CriSel(xxx)
		{
			document.frmReport.txtRepFrom.disabled = true;
			document.frmReport.txtRepTo.disabled = true;
			document.frmReport.cal1.disabled = true;
			document.frmReport.cal2.disabled = true;
			document.frmReport.selInst.disabled = true;
			document.frmReport.selIntr.disabled = true;
			document.frmReport.selTown.disabled = true;
			document.frmReport.txtRepFrom.value = "";
			document.frmReport.txtRepTo.value = "";
			document.frmReport.selInst.value = -1;
			document.frmReport.selIntr.value = -1;
			document.frmReport.selTown.value = -1;
			document.frmReport.selIntrStat.value = 0;
			document.frmReport.selIntrStat.disabled = true;
			if (xxx == 1)
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = false;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
				document.frmReport.selInst.disabled = false;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
			}
			else if (xxx == 2 ||xxx == 9 ||xxx == 11 ||xxx == 12 ||xxx == 13 ||xxx == 14 ||xxx == 15 ||xxx == 6)
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = false;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
				document.frmReport.selInst.disabled = true;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selInst.value = -1;
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
			}
			else if (xxx == 3 ||xxx == 10 ||xxx == 16 ||xxx == 20 ||xxx == 21 ||xxx == 22 || xxx == 19 || xxx == 23 || xxx == 28)
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = false;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
				document.frmReport.selInst.disabled = true;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selInst.value = -1;
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
				if (xxx == 21 || xxx == 22)
				{
					document.frmReport.selIntrStat.disabled = false;

				}
			}
			else if (xxx == 4)
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = false;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
				document.frmReport.selInst.disabled = false;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
			}
			else if (xxx == 5)
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = false;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
				document.frmReport.selInst.disabled = true;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = false;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selInst.value = -1;
				document.frmReport.selIntr.value = -1;
			}
			else if(xxx == 7 || xxx == 18)
			{
				document.frmReport.txtRepFrom.disabled = true;
				document.frmReport.txtRepTo.disabled = true;
				document.frmReport.cal1.disabled = true;
				document.frmReport.cal2.disabled = true;
				document.frmReport.selInst.disabled = true;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selInst.value = -1;
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
			}
			else if(xxx == 8 || xxx == 24)
			{
				document.frmReport.txtRepFrom.disabled = true;
				document.frmReport.txtRepTo.disabled = true;
				document.frmReport.cal1.disabled = true;
				document.frmReport.cal2.disabled = true;
				document.frmReport.selInst.disabled = true;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selInst.value = -1;
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
			}
			else if (xxx == 17 || xxx == 25)
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = true;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = true;
				document.frmReport.selInst.disabled = true;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selInst.value = -1;
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
			}
			else if (xxx == 26 || xxx == 27)
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = true;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = true;
				document.frmReport.selInst.disabled = true;
				document.frmReport.selIntr.disabled = false;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selInst.value = -1;
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
				document.frmReport.selIntrStat.value = 0;
				document.frmReport.selIntrStat.disabled = true;
			}
		}
		function RepGen()
		{
			if (document.frmReport.selRep.value == -1)
			{
				alert("Error: Please select a report type.");
				return;
			}
			if (document.frmReport.chkAddnl.checked == true)
			{
				if (document.frmReport.selLang.value == -1 && document.frmReport.selCli.value == -1 && document.frmReport.selClass.value == -1)
				{
					alert("Error: Please select a filter.");
					return;
				}
			}
			if (document.frmReport.chkAddnl.checked == true)
			{
				if (document.frmReport.selRep.value == 7 || document.frmReport.selRep.value == 8 || document.frmReport.selRep.value == 11 || document.frmReport.selRep.value == 12 || document.frmReport.selRep.value == 13 || document.frmReport.selRep.value == 14 || document.frmReport.selRep.value == 15 || document.frmReport.selRep.value == 17 || document.frmReport.selRep.value == 18 || document.frmReport.selRep.value == 19 || document.frmReport.selRep.value == 6 || document.frmReport.selRep.value == 23 || document.frmReport.selRep.value == 24 || document.frmReport.selRep.value == 26 || document.frmReport.selRep.value == 27)
				{
					alert("Error: Filter is not applicable with this report type.") 
					return;
				}
			}
			if (document.frmReport.selRep.value == 10  || document.frmReport.selRep.value == 19 || document.frmReport.selRep.value == 23)
			{
				if (document.frmReport.txtRepFrom.value == "" || document.frmReport.txtRepTo.value == "")
				{
					alert("Error: Timeframe is required.");
					return;
				}
			}
			if (document.frmReport.selRep.value == 17 || document.frmReport.selRep.value == 25)
			{
				if (document.frmReport.txtRepFrom.value == "")
				{
					alert("Error: Timeframe is required.");
					return;
				}
			}
			if (document.frmReport.selRep.value == 21 || document.frmReport.selRep.value == 22)
			{
									if (document.frmReport.txtRepTo.value == "")
					{
						alert("Error: 'To:' date is required.")
						return; 
					}
					else
					{
						var currentTime = new Date();
						var month = currentTime.getMonth() + 1;
						var day = currentTime.getDate();
						var year = currentTime.getFullYear();
						var datetoday = new Date(month + "/" + day + "/" + year);
						var todate = new Date(document.frmReport.txtRepTo.value);
						var todateyear = todate.getFullYear() + 100;
						var todatemonth = todate.getMonth() + 1;
						var todateday = todate.getDate();
						var newtodate = new Date(todatemonth + "/" + todateday + "/" + todateyear);
 						if (datetoday < newtodate)
 						{
 							alert("Error: 'To:' date should be today or in the past");
 							return;
 						}
 					}
			}
			document.frmReport.action = "action.asp?ctrl=5";
			document.frmReport.submit();
		}
		function PopMe(zzz, xxx)
		{
			if (zzz !== 0)
			{
				newwindow = window.open('printreport.asp','','height=800,width=900,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1');
				if (window.focus) {newwindow.focus()}
			}
		}
		function FilterMe()
		{
			if (document.frmReport.chkAddnl.checked == true)
			{
				document.frmReport.selLang.disabled = false;
				document.frmReport.selCli.disabled = false;
				document.frmReport.selClass.disabled = false;
			}
			else
			{
				document.frmReport.selLang.value = -1;
				document.frmReport.selCli.value = -1;
				document.frmReport.selClass.value = -1;
				document.frmReport.selLang.disabled = true;
				document.frmReport.selCli.disabled = true;
				document.frmReport.selClass.disabled = true;
			}
		}
		function TypeDef(xxx)
		{
			document.frmReport.tadef.value = "";
			if (xxx == 1)
			{
				document.frmReport.tadef.value = "Invoice Report definition here.";
			}
			if (xxx == 8)
			{
				document.frmReport.tadef.value = "List of Interpreters.";
			}
			if (xxx == 7)
			{
				document.frmReport.tadef.value = "List of Requesting Persons";
			}
			if (xxx == 2)
			{
				document.frmReport.tadef.value = "List of Canceled appointments";
			}
			if (xxx == 3)
			{
				document.frmReport.tadef.value = "Bills completed requests for institutions.";
			}
			if (xxx == 4)
			{
				document.frmReport.tadef.value = "per institution Report definition here.";
			}
			if (xxx == 5)
			{
				document.frmReport.tadef.value = "per town Report definition here.";
			}
			if (xxx == 6)
			{
				document.frmReport.tadef.value = "usage Report definition here.";
			}
			if (xxx == 9)
			{
				document.frmReport.tadef.value = "Missed Report definition here.";
			}
			if (xxx == 10)
			{
				document.frmReport.tadef.value = "Language Bank Statistics. NOT YET DONE";
			}
			if (xxx == 11)
			{
				document.frmReport.tadef.value = "Pending requests";
			}
			if (xxx == 12)
			{
				document.frmReport.tadef.value = "Completed requests";
			}
			if (xxx == 13)
			{
				document.frmReport.tadef.value = "Missed requests";
			}
			if (xxx == 14)
			{
				document.frmReport.tadef.value = "Canceled requests";
			}
			if (xxx == 15)
			{
				document.frmReport.tadef.value = "Canceled (Billable) requests";
			}
			if (xxx == 16)
			{
				document.frmReport.tadef.value = "Simulates billing report. This report will not tag requests as billed.";
			}
			if (xxx == 17)
			{
				document.frmReport.tadef.value = "KPI report. Select any date of the month you wish to have a report.";
			}
			if (xxx == 18)
			{
				document.frmReport.tadef.value = "Pending Court requests for the past 30 days.";
			}
			if (xxx == 19)
			{
				document.frmReport.tadef.value = "Completed and Canceled - Billable Court appointments";
			}
			if (xxx == 20)
			{
				document.frmReport.tadef.value = "Audit report.";
			}
			if (xxx == 21)
			{
				document.frmReport.tadef.value = "Pays completed requests for interpreters.";
			}
			if (xxx == 22)
			{
				document.frmReport.tadef.value = "Simulates payroll report. This report will not tag requests as paid.";
			}
			if (xxx == 23)
			{
				document.frmReport.tadef.value = "Cancelled Court Appointments. ";
			}
			if (xxx == 24)
			{
				document.frmReport.tadef.value = "List of ACTIVE interpreters. ";
			}
			if (xxx == 25)
			{
				document.frmReport.tadef.value = "Weekly report. Select any date of the week you wish to have a report.";
			}
			if (xxx == 26)
			{
				document.frmReport.tadef.value = "Mileage report. Select any date of the Month you wish to have a report. You can also select a specific interpreter.";
			}
			if (xxx == 27)
			{
				document.frmReport.tadef.value = "Timesheet report. Select any day of the Week you wish to have a report. You can also select a specific interpreter.";
			}
			if (xxx == 28)
			{
				document.frmReport.tadef.value = "Total Hours report.";
			}
		}
		function CalendarView(strDate)
		{
			document.frmReport.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmReport.submit();
		}
		function SubmitAko()
			{
				document.frmReport.action = 'reqconfirm.asp?ID=' + document.frmReport.hideID.value;
				document.frmReport.submit();
			}
		-->
		</script>
		<body onload='CriSel(document.frmReport.selRep.value); PopMe(<%=tmpRep%>,<%= tmpSele%>);FilterMe();'>
			<form method='post' name='frmReport' action='reports.asp'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" class='bgstyle2' border='0'>
					<tr>
						<td height='100px'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<!-- #include file="_greetme.asp" -->
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td valign='top' >
							<table cellSpacing='4' cellPadding='0' align='center' border='0' bgcolor='#FBEEB7'>
								<tr>
									<td colspan='2' align='center'>
										<b>Report Query</b>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Type:
									</td>
									<td>
										<select class='seltxt' name='selRep'  style='width:200px;' onchange='CriSel(document.frmReport.selRep.value); TypeDef(document.frmReport.selRep.value);'>
											<option value='-1'>&nbsp;</option>
											<% If Request.Cookies("LBUSERTYPE") = 0 Then %>
												<option value='1' <%=TypeSel1%>>Invoice Report</option>
												<option value='8' <%=TypeSel8%>>List - Interpreter</option>
												<option value='24' <%=TypeSel8%>>List - Interpreter (ACTIVE ONLY)</option>
												<option value='7' <%=TypeSel7%>>List - Requesting Person</option>
												<option value='16' <%=TypeSel16%>>Pre-Billing Report</option>
												<option value='22' <%=TypeSel22%>>Pre-Payroll Report</option>
												<option value='4' <%=TypeSel4%>>Per-Institution Report</option>
												<option value='5' <%=TypeSel5%>>Per-Town Report</option>
												<option value='6' <%=TypeSel6%>>Usage Report</option>
												<option value='11' <%=TypeSel11%>>Pending Appointment Report</option>
												<option value='12' <%=TypeSel12%>>Completed Appointment Report</option>
												<option value='13' <%=TypeSel13%>>Missed Appointment Report</option>
												<option value='14' <%=TypeSel14%>>Canceled Appointment Report</option>
												<option value='15' <%=TypeSel15%>>Canceled (Billable) Appointment Report</option>
												<option value='17' <%=TypeSel17%>>KPI Report</option>
												<option value='18' <%=TypeSel18%>>Court Pending Appointment Report</option>
												<option value='19' <%=TypeSel19%>>Court Appointment Report</option>
												<option value='23' <%=TypeSel23%>>Court Cancelled Appointment Report</option>
												<option value='20' <%=TypeSel20%>>Audit Report</option>
											<% ElseIf Request.Cookies("LBUSERTYPE") = 1 Then %>
												<option value='1' <%=TypeSel1%>>Invoice Report</option>
												<option value='8' <%=TypeSel8%>>List - Interpreter</option>
												<option value='24' <%=TypeSel8%>>List - Interpreter (ACTIVE ONLY)</option>
												<option value='7' <%=TypeSel7%>>List - Requesting Person</option>
												<option value='3' <%=TypeSel3%>>Billing Report</option>
												<option value='21' <%=TypeSel21%>>Payroll Report</option>
												<option value='16' <%=TypeSel16%>>Pre-Billing Report</option>
												<option value='22' <%=TypeSel22%>>Pre-Payroll Report</option>
												<option value='4' <%=TypeSel4%>>Per-Institution Report</option>
												<option value='5' <%=TypeSel5%>>Per-Town Report</option>
												<option value='10' <%=TypeSel10%>>Statistics</option>
												<option value='6' <%=TypeSel6%>>Usage Report</option>
												<option value='11' <%=TypeSel11%>>Pending Appointment Report</option>
												<option value='12' <%=TypeSel12%>>Completed Appointment Report</option>
												<option value='13' <%=TypeSel13%>>Missed Appointment Report</option>
												<option value='14' <%=TypeSel14%>>Canceled Appointment Report</option>
												<option value='15' <%=TypeSel15%>>Canceled (Billable) Appointment Report</option>
												<option value='17' <%=TypeSel17%>>KPI Report</option>
												<option value='18' <%=TypeSel18%>>Court Pending Appointment Report</option>
												<option value='19' <%=TypeSel19%>>Court Appointment Report</option>
												<option value='23' <%=TypeSel23%>>Court Cancelled Appointment Report</option>
												<option value='20' <%=TypeSel20%>>Audit Report</option>
												<option value='25' <%=TypeSel25%>>Weekly Report</option>
												<option value='26' <%=TypeSel26%>>Mileage Report</option>
												<option value='27' <%=TypeSel27%>>Timesheet Report</option>
												<option value='28' <%=TypeSel28%>>Total Hours Report</option>
											<% ElseIf Request.Cookies("LBUSERTYPE") = 3 Then %>
												<option value='1' <%=TypeSel1%>>Invoice Report</option>
												<option value='16' <%=TypeSel16%>>Pre-Billing Report</option>
												<option value='22' <%=TypeSel22%>>Pre-Payroll Report</option>
												<option value='3' <%=TypeSel3%>>Billing Report</option>
												<option value='21' <%=TypeSel21%>>Payroll Report</option>
												<option value='17' <%=TypeSel17%>>KPI Report</option>
												<option value='24' <%=TypeSel8%>>List - Interpreter (ACTIVE ONLY)</option>
												<option value='7' <%=TypeSel7%>>List - Requesting Person</option>
												<option value='11' <%=TypeSel11%>>Pending Appointment Report</option>
												<option value='18' <%=TypeSel18%>>Court Pending Appointment Report</option>
											<% End If %>
										</select>
									</td>
								</tr>
								<tr>
									<td align='right' valign='top'>
										Description: 
									</td>
									<td>
										<textarea class='def' name='tadef' readonly ></textarea>
									</td>
								</tr>
								<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
								<tr>
									<td align='right'>
										Criteria:
									</td>
									<td>
										( leave blank to select all )
									</td>
								</tr>
								<tr>
									<td align='right'>Timeframe:</td>
									<td>
										&nbsp;From:<input class='main' size='10' maxlength='10' name='txtRepFrom' readonly value='<%=tmpRepFrom%>'>
										<input type="button" value="..." title='Calendar' name="cal1" style="width: 19px;"
											onclick="showCalendarControl(document.frmReport.txtRepFrom);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
										&nbsp;To:<input class='main' size='10' maxlength='10' name='txtRepTo' readonly value='<%=tmpRepTo%>'>
										<input type="button" value="..." title='Calendar' name="cal2" style="width: 19px;"
											onclick="showCalendarControl(document.frmReport.txtRepTo);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
									</td>
								</tr>
								<tr>
									<td align='right'>
										Institution:
									</td>
									<td>
										<select class='seltxt' name='selInst'  style='width:200px;' onchange=''>
											<option value='0'>&nbsp;</option>
											<%=strInst%>
										</select>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Interpreter:
									</td>
									<td>
										<select class='seltxt' name='selIntr'  style='width:200px;' onchange=''>
											<option value='0'>&nbsp;</option>
											<%=strIntr%>
										</select>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Town:
									</td>
									<td>
										<select class='seltxt' name='selTown'  style='width:200px;' onchange=''>
											<option value='0'>&nbsp;</option>
											<%=strTown%>
										</select>
									</td>
								</tr>
								<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
								<tr>
									<td align='right'>
										Filter:
									</td>
									<td>
										<input type='checkbox' name='chkAddnl' value='1' onclick='FilterMe();'>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Language:
									</td>
									<td>
										<select class='seltxt' name='selLang'  style='width:200px;' onchange=''>
											<option value='0'>&nbsp;</option>
											<%=strLang%>
										</select>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Client:
									</td>
									<td>
										<select class='seltxt' name='selCli'  style='width:200px;' onchange=''>
											<option value='0'>&nbsp;</option>
											<%=strCli%>
										</select>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Classification:
									</td>
									<td>
										<select class='seltxt' style='width: 200px;' name='selClass'>
											<option value='0'>&nbsp;</option>
											<option value='1' <%=SocSer%>>Social Services</option>
											<option value='2' <%=Priv%>>Private</option>
											<option value='3' <%=Court%>>Court</option>
											<option value='4' <%=Med%>>Medical</option>
											<option value='5' <%=legal%>>Legal</option>
										</select>
									</td>
								</tr>
								<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
								<tr>
									<td align='right'><nobr>Interpreter:</td>
									<td>
										<select class='seltxt' style='width: 200px;' name='selIntrStat'>
											<option value='0'>&nbsp;---All---&nbsp;</option>
											<option value='1'>Employee</option>
											<option value='2'>Outside Consultant</option>
										</select>
									</td>
								</tr>
								<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
								<tr>
									<td>&nbsp;</td>
									<td>
										<input class='btn' type='button' style='width: 200px;' value='Generate' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='RepGen();'>
										<input type='hidden' name='hideID'>
									</td>
								</tr>
								<tr>
									<td colspan='2' align='center'>
										<span class='error'><%=Session("MSG")%></span>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
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