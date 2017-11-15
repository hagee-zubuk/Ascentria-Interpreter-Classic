<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	If Z_CZero(Session("UIntr")) > 0 Then
		Set rsAns = Server.CreateObject("ADODB.RecordSet")
		sqlAns = "SELECT DSdoc, initdoc FROM interpreter_T WHERE [index] = " & Session("UIntr")
		rsAns.Open sqlAns, g_strCONN, 1, 3
		If Not rsAns.EOF Then
			rsAns("DSdoc") = Now
			rsAns("initdoc") = Request("txtinit")
			rsAns.Update
		End If
		rsAns.Close
		Set rsAns = Nothing
		Session("MSG") = "Saved."
	End If
End If
If Z_CZero(Session("UIntr")) > 0 Then
	disa = ""
	disa2 = ""
	initdoc = ""
	Set rsDis = Server.CreateObject("ADODB.RecordSet")
	rsDis.Open "SELECT DSdoc, initdoc FROM interpreter_T WHERE [index] = " & Session("UIntr"), g_strCONN, 3, 1
	If Not rsDis.EOF Then
		If Z_FixNull(rsDis("DSdoc")) <> "" Then 
			disa = "disabled"
			disa2 = "disabled checked"
			initdoc = rsDis("initdoc")
		End If
	End If
	rsDis.Close
	Set rsDis = Nothing
End If

<!-- #include file="_closeSQL.asp" -->
%>
<html>
	<head>
		<title>Language Bank - FWA training requirements</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function RTrim(str) {
		  var whitespace = new String(" \t\n\r");
			var s = new String(str);
			if (whitespace.indexOf(s.charAt(s.length-1)) != -1) {
		  	var i = s.length - 1;       
		    while (i >= 0 && whitespace.indexOf(s.charAt(i)) != -1)
		    i--;
		    s = s.substring(0, i+1);
		  }
			return s;
		}
		function LTrim(str) {
		  var whitespace = new String(" \t\n\r");
		  var s = new String(str);
		  if (whitespace.indexOf(s.charAt(0)) != -1) {
		  	var j=0, i = s.length;
		    while (j < i && whitespace.indexOf(s.charAt(j)) != -1)
		    j++;
		    s = s.substring(j, i);
		  }
			return s;
		}
		function Trim(str) {
			return RTrim(LTrim(str));
		}
			function SaveAns() {
				if (document.frmTbl.chkDS.checked == false) {
					alert("Please check the checkbox below.")
					return;
				}
				if (Trim(document.frmTbl.txtinit.value) == "") {
					alert("Please input your intials below.")
					return;
				}
				var ans = window.confirm("Save?\nClick Cancel to stop.");
				if (ans) {
					document.frmTbl.action = "fwa.asp"
					document.frmTbl.submit();
				}
			}
			
		//-->
		</script>
		<body >
			<form method='post' name='frmTbl' action='fwa.asp'>
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
									<td align="center"  colspan='2'>
										<iframe src="files.asp?file=1" width="830" height="600"></iframe>
									</td>
								</tr>
								<tr>
									<td align="center"  colspan='2'>
										<input type="checkbox" name="chkDS" <%=disa2%>> I have read and reviewed the above document.
										<br>
										<input class='main' name='txtinit' size='5' maxlength='4' <%=disa%> value='<%=initdoc%>'>Please input your initials here
										<br>
										<input class='btntbl' type='button' value='Save' <%=disa%> style='height: 25px; width: 150px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='SaveAns();'>
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
										* If your are unable to view the file, you can download the file <a href='#' onclick="document.location='images/FWA-GCT_FINAL.pdf';">HERE</a>.
									</td>
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