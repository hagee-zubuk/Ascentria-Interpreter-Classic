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
	Set rsavail = Server.Createobject("ADODB.RecordSet")
	sqlavail = "SELECT Availability FROM Interpreter_T WHERE [index] = " & Session("UIntr")
	rsavail.Open sqlAvail, g_strCONN, 1, 3
	If not rsavail.EOF Then
		tmpavil = rsavail("Availability")
	End If
	rsavail.Close
	set rsavail = Nothing
%>
<html>
	<head>
		<title>Language Bank - Interpreter Availability</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
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
		function SaveTS()
		{
			document.frmTS.action = "tsheetaction.asp?action=3";
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
								<td class='title' colspan='10' align='center'><nobr> Interpreter Availability</td>
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
										<form name='frmTS' method='POST'>
											<table border='0' cellpadding='1' cellspacing='2' width='75%'>
												<tr>
													<td align='right' width='150px' valign='top'>Availability:</td>
													<td>
														<textarea style='width: 375px;' name='txtAvail' class='main' onkeyup='bawal(this);' ><%=tmpavil%></textarea>
													</td>
												</tr>
											</table>
										
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='12' align='center'>
										*Enter your availabiliy to help our staff determine the proper appointment for you.
									</td>
								</tr>
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