<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
lngID = Session("UIntr")
If lngID < 1 Then
	lngID = CLng(Request("ix"))
	If lngID < 1 Then
		Session("MSG") = "survey response index is missing"
		Response.Redirect "survey.v18.asp"
	End If
End If

Set rsSurv = Server.CreateObject("ADODB.RecordSet")
strSQL = "SELECT [release], [signature] FROM [surveyreports] WHERE [iid]=" & lngID
rsSurv.Open strSQL, g_strCONN, 1, 3
blnRelease = FALSE
dtSig = CDate("2000-01-01")
If Not rsSurv.EOF Then
	blnRelease = CBool( rsSurv("release") )
	dtSig = Z_CDate( rsSurv("signature") )
End If
rsSurv.Close
Set rsSurv = Nothing
If (dtSig > CDate("2018-01-01")) Then
	' it's signed, so it's alright!
	Response.Redirect "survey.report.asp"
End If
If Not blnRelease Then
	Session("MSG") = "survey response index is missing"
	Response.Redirect "survey.v18.asp"
End If
%>
<!doctype html>
<html lang="en">
<head>
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width,initial-scale=1">
	<title>Interpreter Survey</title>
	<meta name="description" content="LanguageBank Internal Interpreter Survey 2018">
	<meta name="author" content="Hagee@zubuk">
 	<link rel="stylesheet" href="css/normalize.css" />
 	<link rel="stylesheet" href="css/skeleton.css" />
 	<link rel="stylesheet" href="css/jquery-ui.min.css" />
	<link rel="stylesheet" href="css/survey.css" />
	<script langauge="javascript" type="text/javascript" src="js/jquery-3.3.1.min.js"></script>
	<script langauge="javascript" type="text/javascript" src="js/jquery-ui.min.js"></script>
  <!--[if lt IE 9]>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html5shiv/3.7.3/html5shiv.js"></script>
  <![endif]-->
	<style>
.ui-autocomplete-loading { background: white url("images/ui-anim_basic_16x16.gif") right center no-repeat; }
	</style>
</head>
<body>
<div class="container">
	<div class="row">
		<div class="twelve columns" id="logobar">
			<img id="logo" src="images/LBISLOGO.jpg" alt="The Language Bank" title="" />
			<h1>Interpreter Performance Evaluation</h1>
		</div>
	</div>

	<div class="row" style="margin-top: 50px;"><div class="one column">&nbsp;</div>
		<div class="ten columns">
			<p>Clicking the "CONTINUE" button below signifies that you electronically sign the Interpreter Evaluation document.</p>
 			<p>This signature does not necessarily indicate agreement, but that you have accepted this performance appraisal. You may respond to it in writing and your comments shall be included your personnel file, along with the apraisal document.</p>			
 			<button class="button button-primary" name="btnContinue" id="btnContinue" value="Continue">Continue</button>
		</div>
	</div>
	<form id="frmGo"" name="frmGo" action="survey.sign.asp" method="post">
		<input type="hidden" id="ixix" name="ixix" value="" />
		<input type="hidden" id="uinr" name="uinr" value="" />
	</form>
</div>
</body>
</html>
<script language="javascript" type="text/javascript"><!--
$( document ).ready(function() {
	$('#btnContinue').click(function(){
		$('#ixix').val('<%=lngID%>');
		$('#uinr').val('<%=lngID%>');
		$('#frmGo').submit();
	});
});
// --></script>