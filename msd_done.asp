<!DOCTYPE html>
<html lang="en">
<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
userid = Z_CLng( Request("fetch") )
strSQL = "SELECT * FROM [Emp_MSD] WHERE [userid]=" & userid
Set rsIR = Server.CreateObject("ADODB.RecordSet")
rsIR.Open strSQL, g_strCONN, 1, 3
If rsIR.EOF Then
	Session("MSG") = "That id was not found. Please re-sign this document."
	Response.Redirect("msd_ia.asp")
End If
rsIR.Close
Set rsIR = Nothing
%>
<head>
	<meta charset="utf-8">
	<title>Authorization for Release of Information</title>
	<meta name="description" content="MSD Interpreter Agreements">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<meta name="author" content="Argao.net">
	<script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"
			integrity="sha256-4+XzXVhsDmqanXGHaHvgh1gMQKX40OUvDEBTu8JcmNs="
			crossorigin="anonymous"></script>
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/normalize/7.0.0/normalize.css"
			integrity="sha256-sxzrkOPuwljiVGWbxViPJ14ZewXLJHFJDn0bv+5hsDY=" crossorigin="anonymous" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/skeleton/2.0.4/skeleton.css"
			integrity="sha256-ECB9bbROLGm8wOoEbHcHRxlHgzGqYpDtNTgDTyDz0wg=" crossorigin="anonymous" />
<style type="text/css">
h1 { font-size: 1.6em; }
h2 { font-size: 1.4em; font-weight: bold; }
</style>			
</head>
<body>
	<div class="container">
	<div class="row" style="text-align: center;">
		<img src='images/LBISLOGO.jpg' border="0" style="width: 287px; height: 64px;" />
		<div class="twelve columns"><h1>Manchester School District Confidentiality and<br />
		Data Security Agreements for Language Bank Interpreters</h1></div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="ten columns">
			<h1>Thank you</h1>
			<p>You may close this window</p>
		</div>
	</div>
	</div> <!-- container -->
</body>
</html>
