<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
Function GetEmail(xxx)
	GetEmail = ""
	Set rsEm = Server.CreateObject("ADODB.RecordSet")
	sqlEm = "SELECT [e-mail] FROM interpreter_T WHERE [index] = " & xxx
	rsEm.Open sqlEm, g_strCONN, 1, 3
	If Not rsEm.EOF Then
		GetEmail = rsEm("e-mail")
	End If
	rsEm.Close
	Set rsEm = Nothing
End Function
If Request("mail") = 1 Then 'Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	'GET EMAIL INFO
	Set rsReq = Server.CreateObject("ADODB.REcordSet")
	sqlReq = "SELECT * FROM request_T WHERE [index]= " & Request("ID")
	rsReq.Open sqlReq, g_strCONN, 1, 3
	If Not rsReq.EOF Then
		appdate = rsReq("appdate")
		apptime = CTime(rsReq("apptimefrom")) & " - " & CTime(rsReq("apptimeto"))
		appCity = GetCity(rsReq("DeptID"))
		If rsReq("cliadd") = 1 Then appCity = rsReq("Ccity")
		IntrLang = Ucase(GetLang(rsReq("langID")))
	End If
	rsReq.Close
	Set rsReq = Nothing
	'SEND EMAIL
	'on error resume next
	Set mlMail = CreateObject("CDO.Message")
With mlMail.Configuration
	.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")		= 2
	.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")		= "smtp.socketlabs.com"
	.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")	= 2525
	.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername")		= "server3874"
	.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")		= "UO2CUSxat9ZmzYD7jkTB"
	.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")	= 1 'basic (clear-text) authentication
	.Fields.Update
End With
	mlMail.To = GetEmail(Request("selIntr"))
	mlMail.BCC = "patrick@zubuk.com"
	mlMail.From = "language.services@thelanguagebank.org"
	mlMail.Subject= "Appointment on " & appDate & " at " & apptime & " in " & appCity
	strMSG = "Are you available to do this appointment?<br><br>"& _
		"If you accept this appointment this is the amount you will be reimbursed for mileage and travel time. " & _
		"Payable travel time is " & Z_Czero(Request("txtTravel")) & " hrs. " & _
		"and payable mileage is " & Z_Czero(Request("txtMile")) & "  miles.<br><br>"& _
		"Please reply to this email or contact " & Request.Cookies("LBUsrName") & " of LanguageBank.<br><br>"& _
		"Thank you.<BR><BR><BR>" & _
		"<font color='#FFFFFF'>" & Request("adr1") & "|" & Request("adr2") & "|" & Request("zip1") & "|" & Request("zip2") & "</font>"
	mlMail.HTMLBody = "<html><body>" & vbCrLf & strMSG & vbCrLf & "</body></html>"
	mlMail.Send
	set mlMail=nothing
	'save to notes
	IntrName = GetIntr2(Request("selIntr"))
	Set rsNotes = Server.CreateObject("ADODB.RecordSet")
	sqlNotes = "SELECT LBComment FROM request_T WHERE [index] = " & Request("ReqID")
	rsNotes.Open sqlNotes, g_StrCONN, 1, 3
	If Not rsNotes.EOF Then
		rsNotes("LBComment") = rsNotes("LBComment") & vbCrlF & "Email sent to " & IntrName & " on " & now & " for availability"
		rsNotes.Update
	End If
	rsNotes.CLose
	Set rsNotes = Nothing
	Session("MSG") = "Email Sent."
End If
	'PREPARE EMAIL	
	Set rsReq = Server.CreateObject("ADODB.REcordSet")
	sqlReq = "SELECT * FROM request_T WHERE [index] = " & Request("ID")
	rsReq.Open sqlReq, g_strCONN, 1, 3
	If Not rsReq.EOF Then
		appdate = rsReq("appdate")
		apptime = rsReq("apptimefrom") & " - " & rsReq("apptimeto")
		appCity = GetCity(rsReq("DeptID"))
		If rsReq("cliadd") = 1 Then appCity = rsReq("Ccity")
		IntrLang = Ucase(GetLang(rsReq("langID")))
	End If
	rsReq.Close
	Set rsReq = Nothing
	strSubj = "Appointment on " & appDate & " at " & apptime & " in " & appCity
	strMSG = "Are you available to do this appointment?" & vbCrlf & vbCrlf & _
		"Please reply to this email or contact " & Request.Cookies("LBUsrName") & " of LanguageBank." & vbCrlf & vbCrlf & _
		"If you accept this appointment this is the amount you will be reimbursed for mileage and travel time." & vbCrlf & _
		"Payable travel time is " & Z_Czero(Request("txtTravel")) & " hrs. and payable mileage is " & Z_Czero(Request("txtMile")) & " miles." & vbCrlf & vbCrlf & _
		"Thank you."
	'INTERPRETER LIST
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT * FROM interpreter_T WHERE (Upper(Language1) = '" & IntrLang & "' OR Upper(Language2) = '" & IntrLang & "' OR Upper(Language3) = '" & IntrLang & _
		"' OR Upper(Language4) = '" & IntrLang & "' OR Upper(Language5) = '" & IntrLang & "') AND Active = 1 AND [e-mail] <> '' ORDER BY [Last Name], [First Name]"
	rsIntr.Open sqlIntr, g_strCONN, 1, 3
	Do Until rsIntr.EOF
		'include vacation
		
		IntrName = rsIntr("Last Name") & ", " & rsIntr("First Name")
		If isNull(rsIntr("vacfrom")) Then
			myIntr = ""
			If  cint(Request("selIntr")) = rsIntr("index") Then myIntr = "selected"
			strIntr = strIntr & "<option value='" & rsIntr("index") & "' " & myIntr & ">" & IntrName & "</option>" & vbCrlf
		Else
			If Not (appdate >= rsIntr("vacfrom") And appdate <= rsIntr("vacto")) Then
				myIntr = ""
				If  cint(Request("selIntr")) = rsIntr("index") Then myIntr = "selected"
				strIntr = strIntr & "<option value='" & rsIntr("index") & "' " & myIntr & ">" & IntrName & "</option>" & vbCrlf
			End If
		End If
		rsIntr.MoveNext
	Loop
	rsIntr.Close
	Set rsIntr = Nothing
'End If
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Email Interpreter</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
			function sendMe()
			{
				if (document.frmEmail.selIntr.value == 0)
				{
					alert("ERROR: Please select an interpreter.")
					return;
				}
				else
				{
					document.frmEmail.mail.value = 1;
					document.frmEmail.action = "emailIntr.asp";
					document.frmEmail.submit();	
				}
			}
			function GetMile(xxx)
			{
				document.frmEmail.action = "Intrmile.asp?selIntr=" + xxx + "&RID=" + <%=Request("ID")%>;
				document.frmEmail.submit();	
			}
		</script>
	</head>
	<body>
		<form name='frmEmail' method='post'>
			<table cellpadding='1' cellspacing='0' border='0'>
				<tr><td>&nbsp;
					<input name="mail" value="" type="hidden" />
					</td></tr>
				<tr>
					<td align='right'>&nbsp;Interpreter:</td>
					<td align='left'>&nbsp;
						<select class='seltxt' name='selIntr' style='width: 150px;' onchange='GetMile(this.value);'>
							<option value='0'>&nbsp;</option>
							<%=strIntr%>
						</select>
					</td>
				</tr>
				<tr>
					<td align='right'>Subject:</td>
					<td align='left'>&nbsp;<b><%=strSubj%></b></td>
				</tr>
				<tr>
					<td align='right' valign='top'>Message:</td>
					<td align='left'>&nbsp;
						<textarea name='txtMSG' readonly cols='48' rows='10'>
							<%=strMSG%>
						</textarea>
					</td>
				</tr>
				<tr>
					<td align='right' colspan='2'>
						<input class='btn' type='button' value='Send' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='sendMe();'>
						<input class='btn' type='button' value='Close' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='window.close();'>
						<input type='hidden' name='ReqID' value='<%=Request("ID")%>'>
						<input type="hidden" name="id" value='<%=Request("ID")%>' />
						<input type="hidden" name="txtTravel" value='<%=Request("txtTravel")%>' />
						<input type="hidden" name="txtMile" value='<%=Request("txtMile")%>' />
						<input type='hidden' name='adr1'  value='<%=Request("adr1")%>'>
						<input type='hidden' name='adr2'  value='<%=Request("adr2")%>'>
						<input type='hidden' name='zip1'  value='<%=Request("zip1")%>'>
						<input type='hidden' name='zip2'  value='<%=Request("zip2")%>'>
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