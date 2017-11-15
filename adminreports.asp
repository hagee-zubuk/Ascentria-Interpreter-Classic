<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Security.asp" -->
<%
tmpPage = "document.frmReport."
%>
<html>
	<head>
		<title>Language Bank - Admin Reports</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
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
		function RepGen(xxx,yyy, tmpfrom, tmpto)
		{
			if (yyy == '')
			{
				yyy = 0;
			}
			newwindow = window.open('Intrreports.asp?ctrl=2&selRep=' + xxx + '&txtyear=' + yyy + '&txtRepFrom=' + tmpfrom + '&txtRepTo=' + tmpto + '','','height=800,width=900,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1');
				if (window.focus) {newwindow.focus()}
			//document.frmReport.action = "Intrreports.asp?ctrl=2"
			//document.frmReport.submit();
		}
		function CriSel(xxx)
		{
			document.frmReport.txtyear.disabled = true;
			document.frmReport.cal1.disabled = true;
			document.frmReport.cal2.disabled = true;
			if (xxx == 1 || xxx == 2)
			{
				document.frmReport.txtyear.disabled = false;
				document.frmReport.cal1.disabled = true;
				document.frmReport.cal2.disabled = true;
			}
			if (xxx == 4 || xxx == 6)
			{
				document.frmReport.txtyear.disabled = true;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
			} 
		}
		function CalendarView(strDate)
			{
				document.frmReport.action = 'calendarview2.asp?appDate=' + strDate;
				document.frmReport.submit();
			}
		</script>
	</head>
	<body onload='CriSel(0);'>
		<form method='post' name='frmReport' action='Intreports.asp?ctrl=2'>
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
										<b>Admin Report Query</b>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Type:
									</td>
									<td>
										<select class='seltxt' name='selRep'  style='width:200px;' onchange='CriSel(document.frmReport.selRep.value);'>
											<option value='0'>&nbsp;</option>
											<option value='1' <%=TypeSel1%>>Interpreter Training</option>
											<option value='2' <%=TypeSel2%>>Interpreter Evaluation/Feedback</option>
											<option value='3' <%=TypeSel3%>>Interpreter Documents</option>
											<option value='4' <%=TypeSel4%>>Interpreter Date of Hire</option>
											<option value='6' <%=TypeSel6%>>Interpreter Date of Termination</option>
											<option value='5' <%=TypeSel5%>>Interpreter Driving and Criminal Record</option>
										</select>
									</td>
								</tr>
								<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
								<tr>
									<td align='right'>
										Year:
									</td>
									<td>
										<input class='main' size='5' maxlength='4' name='txtyear' value='' onkeyup='bawal(this);'>
										<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">yyyy</span>
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
								<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
								<tr>
									<td>&nbsp;</td>
									<td>
										<input class='btn' type='button' style='width: 200px;' value='Generate' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='RepGen(document.frmReport.selRep.value, document.frmReport.txtyear.value, document.frmReport.txtRepFrom.value, document.frmReport.txtRepTo.value);'>
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