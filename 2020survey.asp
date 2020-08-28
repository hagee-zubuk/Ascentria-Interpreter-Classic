<!doctype html>
<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
Function Z_YMDDate(dtDate)
DIM lngTmp, strDay, strTmp
	If Not IsDate(dtDate) Then Exit Function
	strTmp = DatePart("yyyy", dtDate)
	Z_YMDDate = strTmp & "-"

	lngTmp = CLng(DatePart("m", dtDate))
	If lngTmp < 10 Then Z_YMDDate = Z_YMDDate & "0"
	Z_YMDDate = Z_YMDDate & lngTmp & "-"
	lngTmp = CLng(DatePart("d", dtDate))
	If lngTmp < 10 Then Z_YMDDate = Z_YMDDate & "0"
	Z_YMDDate = Z_YMDDate & lngTmp
End Function


'check if interpreter has filled something up for today
IntrID = Session("UIntr")
strSQL = "SELECT COUNT([id]) AS c FROM [2020Survey] WHERE [IntrID]=" & IntrID & " AND [SvyDt]='" & Z_YMDDate(Date) & "'"
Set rsZ = Server.CreateObject("ADODB.Recordset")
rsZ.Open strSQL, g_strCONNDB, 3, 1
If rsZ("C") > 0 Then
	rsZ.Close
	Set rsZ = Nothing
	Response.Redirect "calendarview2.asp"
End If
rsZ.Close
Set rsZ = Nothing

%>
<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta name="viewport" content="width=device-width, initial-scale=1.0" />
	<meta name="description" content="COVID-19 Update page" />
	<meta name="author" content="Zubuk PH" />

	<title>Interpreter Survey :: Language Bank</title>
	<link type="text/css" rel="stylesheet" href="style.css" />
	<link type="text/css" rel="stylesheet" href="CalendarControl.css" />
	<script src="CalendarControl.js" language="javascript"></script>
	<script src="js/jquery-3.3.1.min.js" language="javascript"></script>
	<style>
button#btnOK { width: 150px; height: 30px; }
form#frmInt { display: none; }
div.container { margin: 0px 20px; }
div.ynblock { display: inline-block; height: 45px; width: 15%; min-width: 120px; text-align: left; float: left; }
div.blurb   { display: inline-block; height: 45px; width: 80%; min-width: 420px; text-align: left; float: left; padding: 0px 10px;}
div.question { clear: both; width: 100%; margin-top: 20px; padding-top: 10px; }
p.blurb { display: inline-block; font-size: 110%; background-color: #D6EAF8; }
input[type="text"] { text-align: center; height: 20px; padding: 2px 10px;}
	</style>
</head>
<body>
<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
	<tr><td valign='top'><!-- #include file="_header.asp" --></td></tr>
</table>
<div class="container">
<h1>Language Bank Interpreter Survey</h1>
<p>Please complete this survey before continuing on to the site.</p>
	<form autocomplete="off" id="frmInt" method="POST" name="frmInt" style="max-width: 640px;">
		<!-- changes on 2020-06-05:


1.    Have you traveled out of state within last 14 days?

2.    Have you come in close contact with someone (within 6 feet) who has tested positive for COVID-19 or is suspected of having COVID-19? This includes members of your household as well as clients/patients.

3.    Do you have a fever (greater than 100.4F or 38.0C) OR symptoms of lower respiratory illness such as cough, shortness of breath, OR difficulty breathing?
-->
	<div class="question" id="q1">
		<!-- h3>Have you travelled internationally in the past 14 days?</h3 -->
		<h3>Have you traveled out of New Hampshire, Massachusetts, Maine, Vermont and Rhode Island within last 14 days?</h3>
		<div class="ynblock">
		<input type="radio" name="q1" id="q1_yes" value="1" />&nbsp;YES<br />
		<input type="radio" name="q1" id="q1_no"  value="0" />&nbsp;No<br />
		</div>
	</div>

	<div class="question" id="q2">
		<!-- h3>Have you come in close contact with someone who has a laboratory-confirmed case of COVID-19, aka the coronavirus?</h3 -->
		<h3>Have you come in close contact (within 6 feet) with someone who has tested positive
			for COVID-19 or is suspected of having COVID-19? This includes members of your
			household as well as clients/patients.</h3>
		<div class="ynblock">
			<input type="radio" name="q2" id="q2_yes" value="1" />&nbsp;YES<br />
			<input type="radio" name="q2" id="q2_no"  value="0" />&nbsp;No<br />
		</div>
		<div class="blurb">
			<p class="blurb" id="q2b">It is important that you isolate yourself and call your doctor immediately.
				A Language Bank employee will reach out to you as soon as possible.
			</p>
		</div>
	</div>

	<div class="question" id="q3">
		<!-- h3>Do you have a fever (greater than 100.4&deg;F or 38.0&deg;C) OR symptoms of lower respiratory illness such as cough, shortness of breath, OR difficulty breathing?</h3 -->
		<h3>Do you have a fever (greater than 100.4&deg;F or 38.0&deg;C) OR symptoms of
			lower respiratory illness such as cough, shortness of breath, OR difficulty
			breathing?</h3>
		<div class="ynblock">
			<input type="radio" name="q3" id="q3_yes" value="1" />&nbsp;YES<br />
			<input type="radio" name="q3" id="q3_no"  value="0" />&nbsp;No<br />
		</div>
		<div class="blurb">
			<p class="blurb" id="q3b">It is important that you isolate yourself and call your doctor immediately.
				Please do not accept any appointments or perform any work for Language Bank at this time.
			</p>
		</div>
	</div>

	<div class="question" id="q4" style="margin-top: 100px;">
		<h3>Enter your initials below</h3>
		<p>By typing my initials in the space below, I agree to notify Language Bank if the responses I have provided 
			to the questions on this form change at any time. Particularly, if I become ill, have a fever, or travel
			internationally after completing this survey I will notify Language Bank immediately.
			</p>
		<div style="text-align: center; width: 100%;">
			<input id="txtSig" name="txtSig" type="text" maxlength="6" value="" autocomplete="" placeholder="initial here">
		</div>
	</div>
	<div class="question" style="text-align: center; width: 100%;" >
		<button type="button" name="btnOK" id="btnOK" class="btn button-primary">Continue</button>
	</div>
	</form>
</div><!-- container -->

</body>
</html>
<script>
function chkFormInput() {
	var blnSub = true;
	if (!$("input[name='q1']:checked").val()) {
		console.log('Q1: Nothing is checked!');
		blnSub = false;
	}
	if (!$("input[name='q2']:checked").val()) {
		console.log('Q2: Nothing is checked!');
		blnSub = false;
	}
	if (!$("input[name='q3']:checked").val()) {
		console.log('Q3: Nothing is checked!');
		blnSub = false;
	}
	var txtSig = $('#txtSig').val();
	txtSig = txtSig.trim();
	if (txtSig.length < 2) {
		console.log('Sig: not set!');
		blnSub = false;
	}

	if (blnSub) {
		// submit the survey form
		console.log('All prerequisites complete. Sending!');
		$("#frmInt").attr('action', '2020survey_send.asp');
	}
	return blnSub;
}

$(document).ready(function() {
	$('#q2_yes').click(function() {
			if($('#q2_yes').prop("checked")) { $('#q2b').show(); } else { $('#q2b').hide(); }
		});
	$('#q2_no').click(function() {
			if($('#q2_yes').prop("checked")) { $('#q2b').show(); } else { $('#q2b').hide(); }
		});
	$('#q3_yes').click(function() {
			if($('#q3_yes').prop("checked")) { $('#q3b').show(); } else { $('#q3b').hide(); }
		});
	$('#q3_no').click(function() {
			if($('#q3_yes').prop("checked")) { $('#q3b').show(); } else { $('#q3b').hide(); }
		});
	$('p.blurb').hide();
	$('#btnOK').click(function() {
			//
			$('#frmInt').submit();
		});
	$('#txtSig').val('');
	
	$('form#frmInt').submit(function(event){
		if (chkFormInput()) {
			$('form#frmInt').hide(1000);
		    return;
  		} else {
 			alert("Please complete the form\nand type your initials in the space provided");
  			event.preventDefault();
  		}
	});

	$('form#frmInt').show();
});
</script>