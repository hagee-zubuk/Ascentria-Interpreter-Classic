<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) <> 1 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
Function CleanFax(strFax)
	CleanFax = Replace(strFax, "-", "") 
End Function
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	server.scripttimeout = 360000
	Set rsEmail = Server.CreateObject("ADODB.RecordSet")
	sqlEmail = "SELECT Email, Fax, index from requester_T ORDER BY index"
	rsEmail.Open sqlEmail, g_strCONN, 1, 3
	Do Until rsEmail.EOF
		If rsEmail("email") <> "" Or rsEmail("fax") <> "" Then
			myEmailAdr = rsEmail("email")
			If myEmailAdr = "" Then myEmailAdr = CleanFax(rsEmail("fax")) & "@emailfaxservice.com" 

			retVal = zSendMessage(myEmailAdr, "", Request("txtSub"), Request("txtMSG"))

		End If
		rsEmail.MoveNext
	Loop
	rsEmail.Close
	Set rsEmail = Nothing
	Session("MSG") = "Email Sent."
	'response.redirect "main.asp"
End If
%>
<html>
	<head>
		<title>LanguageBank - Mass Email</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function SendMe()
		{
			if (document.frmMemail.txtSub.value == "")
			{
				alert("Please include a subject.")
				return;
			}
			if (document.frmMemail.txtMSG.value == "")
			{
				alert("Please include a message.")
				return;
			}
			var ans = window.confirm("Send eMail? This might take a few minutes to complete.");
			if (ans){
			document.frmMemail.action = "massmail.asp";
			document.frmMemail.submit();
			}
		}
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
		-->
		</script>
	</head>
	<body>
		<form method='post' name='frmMemail'>
			<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
				<tr>
					<td valign='top'>
						<!-- #include file="_header.asp" -->
					</td>
				</tr>
				<tr>
					<td valign='top'>
						<table cellSpacing='2' cellPadding='0' width="100%" border='0'>
							<!-- #include file="_greetme.asp" -->
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td align='center' colspan='10'>
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
								<td align='center'>
									<table border='0' cellspacing='1'>
											<tr>
											<td valign='top'>Subject:</td>
											<td>
												<input class='main' size='50' maxlength='50' name='txtSub' onkeyup='bawal(this);'>
											</td>
										</tr>
										<tr>
											<td valign='top'>Message:</td>
											<td>
												<textarea class='main' style='width: 400px;' name='txtMSG' cols='75' rows='10' onkeyup='bawal(this);'>
												</textarea>
											</td>
										</tr>
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td colspan='2' align='right'>
												<input class='btn' type='button' value='Send' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SendMe();'>
											</td>
										</tr>
									</table>
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
</html>
<%
If Session("MSG") <> "" Then
	tmpMSG = Session("MSG")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>
