<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Security.asp" -->
<%
download = 0
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	rsUser.Open "SELECT password from User_T WHERE [index] = " & Session("UsrID"), g_strCONN, 3, 1
	If Not rsUser.EOF Then
		If Z_DoDecrypt(rsUser("password")) = Request("txtPW") Then
			rsUser.Close
			Set rsUser = Nothing
			If Z_GetInfoFROMAppID(Request("ReqID"), "InstID") = 860 Then ' UMass
				Response.Redirect "UMass.asp?reqid=" & Request("reqid")
			Else
				Response.Redirect "close.asp?reqid=" & Request("reqid")
			End If
		Else
			Session("MSG") = "Invalid password."
		End If
	End If
	rsUser.Close
	Set rsUser = Nothing
End If
%>
<html>
	<head>
		<title>Language Bank - Download Verification Form</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function dloadme(xxx)
		{
			if (document.frmVer.txtPW.value == '') {
				alert("Password is blank.");
				return;
			}
			else {
				document.frmVer.action = 'dloadvform.asp?reqid=' + xxx;
				document.frmVer.submit();
			}
		}
		-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form method='post' name='frmVer'>
			<table cellpadding='0' cellspacing='0' border='0' align='left' height='95%' width='100%'>
				<tr>
					<td class='header' colspan='2'><nobr>Verification  --&gt&gt</td>
				</tr>
				<tr>
					<td align='right'>
						Password:
					</td>
					<td><input type='password' class='main' style='width: 130px;' maxlength='20' name='txtPW'></td>
				</tr>
				<tr>
					<td colspan='2' align='center'>
						<input class='btn' type='button' value='Verify' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='dloadme(<%=Request("reqid")%>);'>
					</td>
				</tr>
				<tr>
					<td colspan='3' align='right' valign='bottom'>
						<font size='1'><i><u>* Password must match the password of the current user.</u></i></font>
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