<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
<!-- #include file="_closeSQL.asp" -->
%>
<html>
	<head>
		<title>Language Bank - Trainings &amp; Links</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		//-->
		</script>
<style>
.trainings p {
	margin-bottom: 20px;
}
.trainings p a:active, .trainings p a:link, .trainings p a:visited {
	color: blue;
	font-size: 110%;
	font-weight: bold;
}
.trainings p a:hover {
	background-color: yellow;
	color: blue;
	font-size: 110%;
	font-weight: bold;
}
</style>		
</head>		
		<body >
			<form method='post' name='frmTbl' action='#'>
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
								<tr><td>&nbsp;</td></tr>
								<tr><td>
									<div style="margin-left: 50px;">
									<h1>TRAININGS AND LINKS</h1>
										<div style="margin-left: 20px;" class="trainings">
	<p><a class="Admin" href="http://intranet.ascentria.org/default.aspx" target="_blank">Ascentria Intranet</a></p>
	<p><a class="Admin" href="https://securefile.ascentria.org/filedrop/LBInterpreterInfo" target="_blank">Ascentria Secure File Transfer</a><br />
	Please use this link to upload your credentials (training and cont. education certificates) and proof of immunization ONLY</p>
	<p><a class="Admin" href="fwa.asp" target="_self">FWA Training</a></p>
	<p><a class="Admin" href="http://www.imiaweb.org/code/" target="_blank">IMIA Code of Ethics for Medical Interpreters</a></p>
	<p><a class="Admin" href="pdf.asp" target="_self">Interpreter Guidlines</a></p>
	<p><a class="Admin" href="http://www.najit.org/about/NAJITCodeofEthicsFINAL.pdf" target="_blank">NAJIT Code of Ethics for Judicial Interpreters</a></p>
	<p><a class="Admin" href="http://www.courts.state.nh.us/supreme/orders/12-24-13-order-appendix-b.pdf" target="_blank">NHJB Code of Professional Responsibility for Interpreters</a></p>
	<p><a class="Admin" href="http://lssne.training.reliaslearning.com/lib/Authenticate.aspx?ReturnUrl=%2f" target="_blank">Relias Learning /Training</a></p>
	<p><a class="Admin" href="https://ew42.ultipro.com/Login.aspx?ReturnUrl=%2fdefault.aspx  " target="_blank">UltiPro</a></p>
										</div>
									</div>
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