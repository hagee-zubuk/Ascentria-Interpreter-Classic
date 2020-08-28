<!DOCTYPE html>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
chksig	= Request("chkSig")
fname	= Request("fname")
mname	= Request("mname")
lname	= Request("lname")
suffix	= Request("suffix")
strLAst = ""
strIP	= Request.ServerVariables("REMOTE_ADDR")
strUA	= Request.ServerVariables("HTTP_USER_AGENT")
blnDisp = FALSE

fetchid = Z_CLng(Request("fetch"))
Set rsUser = Server.CreateObject("ADODB.RecordSet")
If fetchid <= 0 Then
	strSQL = "SELECT u.[Index] AS UserID, u.[Fname], u.[Lname], i.[address1], i.[address2], i.[city], i.[state], i.[zip code] " & _
			", i.[e-mail], i.[Phone1] " & _
			"FROM [user_t] AS u LEFT JOIN [interpreter_T] as i ON u.[IntrID]=i.[index] " & _
			"WHERE u.[username]='" & Request.Cookies("LBUSER") & "'"
	rsUser.Open strSQL, g_strCONN, 3, 1
	If rsUser.EOF Then
		strEmpName = ""
		strAddr = ""
		strCellno = ""
		strEmail = ""
		lngUserID = 0
	Else
		strEmpName = Z_FixNull(rsUser("Lname"))
		If strEmpName <> "" Then strEmpName = " " & strEmpName
		strEmpName = Z_FixNull(rsUser("Fname")) & strEmpName
		strState = Z_FixNull(rsUser("state"))
		strCity = Z_FixNull(rsUser("city"))
		If strCity <> "" And strState <> "" Then strCity = strCity & ", " & rsUser("state") & " " & rsUser("zip code")
		strAddr = Z_FixNull(rsUser("address2"))
		If strAddr <> "" Then strAddr = ", " & strAddr
		strAddr = Z_FixNull(rsUser("address1")) & strAddr
		If strCity <> "" Then strAddr = strAddr & ", " & strCity
		strCellno = Z_FixNull(rsUser("Phone1"))
		strEmail = Z_FixNull(rsUser("e-mail"))
		lngUserID = Z_CLng(rsUser("UserID"))
	End If
Else
	strSQL = "SELECT [userid], [empname], [addr], [cellno], [email], [fname], [mname], [lname], [suffix], [chksig], [ip], [useragent], [last] " & _
			"FROM [ese_form] WHERE [userid]=" & fetchid
	rsUser.Open strSQL, g_strCONN, 3, 1
	If rsUser.EOF Then
		strEmpName = ""
		strAddr = ""
		strCellno = ""
		strEmail = ""
		lngUserID = 0
	Else
		blnDisp = True
		strEmpName = Z_FixNull(rsUser("empname"))
		strAddr = Z_FixNull(rsUser("addr"))
		strCellno = Z_FixNull(rsUser("cellno"))
		strEmail = Z_FixNull(rsUser("email"))
		lngUserID = Z_CLng(rsUser("userid"))
		fname	= Z_FixNull(rsUser("fname"))
		mname	= Z_FixNull(rsUser("mname"))
		lname	= Z_FixNull(rsUser("lname"))
		suffix	= Z_FixNull(rsUser("suffix"))
		strIP	= Z_FixNull(rsUser("ip"))
		strUA	= Z_FixNull(rsUser("useragent"))
		strLast = FormatDateTime(rsUser("last"), 0)
		If rsUser("chksig") = 1 Then chksig = "checked"
	End If
End If
rsUser.Close
Set rsUser = Nothing
%>
<html lang="en">
<head>
	<meta charset="utf-8">
	<title>Employee Standards & Expectations</title>
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<meta name="author" content="Argao.net">
<%
If (blnDisp) Then
	'meh
Else
%>
	<script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"
			integrity="sha256-4+XzXVhsDmqanXGHaHvgh1gMQKX40OUvDEBTu8JcmNs="
			crossorigin="anonymous"></script>
	<script src="foi.js"></script>
<%
End If
%>
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/normalize/7.0.0/normalize.css"
			integrity="sha256-sxzrkOPuwljiVGWbxViPJ14ZewXLJHFJDn0bv+5hsDY=" crossorigin="anonymous" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/skeleton/2.0.4/skeleton.css"
			integrity="sha256-ECB9bbROLGm8wOoEbHcHRxlHgzGqYpDtNTgDTyDz0wg=" crossorigin="anonymous" />
<style type="text/css">
h1 { font-size: 1.6em; }
</style>						
</head>
<body>
	<div class="container">
	<div class="row" style="text-align: center;">
		<img src='images/LBISLOGO.jpg' border="0" style="width: 287px; height: 64px;" />
		<div class="twelve columns"><h1>Employee Standards and Expectations to Maintain a Safe
		and Supportive Workplace During Phase 1 of the Coronavirus Pandemic</h1></div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="ten columns">
	<p>The coronavirus is significantly impacting all of us and requiring that we rethink and reconsider how we live and interact in consideration of our own health and everyone around us. In order to protect the Ascentria community, both employees and the people we serve, from the spread of the coronavirus we are implementing the following practices recommended by the Center for Disease Control-CDC. Your signature on this document acknowledges that you understand and will follow the expectations outlined below:</p>
		</div>
	</div>
	<div class="row">
		<div class="two columns">&nbsp;</div>
		<div class="eight columns">
	<ul>
		<li>Wear a facemask whenever I am working for Ascentria outside of my home.</li>
		<li>Regularly clean my cloth facemask (at least every two days) to ensure it
				functions properly, and/or replace it when it no longer fits properly
				or becomes dirty.</li>
		<li>Actively maintain social distancing of at least 6 feet from others. I
				understand that small examination rooms may make this difficult at
				times, but will do my best to abide by this requirement.</li>
		<li>Clean my hands thoroughly, regularly and often when working. Washing is
				the preferred method.</li>
		<li>Agree to stay home if I have symptoms consistent with the coronavirus
				and let my manager know if I have symptoms or if I have had close
				contact with someone with a confirmed case.</li>
		<li>Discuss my work schedule with my manager, and adjust as necessary, to
				limit the number of staff congregating at one time.</li>
		<li>Other than interpretations, I will not schedule or attend in person
				meetings, either internal Ascentria meetings or external meetings
				unless my Executive Team leader has approved an in-person meeting.</li>
		<li>Even when I am not working, I will observe social distancing practices,
				wear a mask when in contact with people outside of my home, and
				follow government mandates to keep me and the people I work with
				safe.</li>
		<li>If I have a health condition that puts me at heightened risk for COVID-
				19, I will contact Human Resources to see if I am eligible for a
				Family Medical Leave.</li>
		<li>If I have or develop a health condition that does not allow me to
				comply with this agreement, I will contact human resources immediately
				for guidance.</li>
		<li>I will respectfully communicate with my manager and colleagues at all
				times understanding that these expectations are new.</li>
		<li>I agree to have my temperature taken and complete a health assessment,
				if required by my manager, or the institution where I am
				interpreting.</li>
	</ul>
		</div>
	</div>

	<form id="frmROI" name="frmROI" method="post" action="ese_proc.asp">
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="six columns">
			<label for="empName">Name of Employee:</label>
			<input class="u-full-width" type="text" placeholder="Employee Name" id="empname" name="empname"
					value="<%= strEmpName%>" readonly="readonly"
					/>
			<input type="hidden" name="userid" id="userid" value="<%= lngUserID %>" />
		</div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="ten columns">
			<label for="addr">Employee Address:</label>
			<input class="u-full-width" type="text" placeholder="Employee Address" id="addr" name="addr" 
					value="<%= strAddr%>" />
		</div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="four columns">
			<label for="addr">Cell Phone Number:</label>
			<input class="u-full-width" type="tel" maxlength="20" placeholder="Cell #" id="cellno" name="cellno"
					value="<%= strCellno%>" />
		</div>
		<div class="six columns">
			<label for="addr">e-Mail Address:</label>
			<input class="u-full-width" placeholder="e-mail address" type="email" id="email" name="email"
					value="<%= strEmail%>" />
		</div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="ten columns">
	<div id="sigsection" style="border: 1px dotted #777; padding: 5px;">
				<u style="font-size: 140%;">Signature</u>
<div id="sig_1">
	<p style="font-weight: bold;">Step 1: Check the box below</p>
	<p style="text-align: center;"><input type="checkbox" value="1" id="chkSig" name="chkSig" <%=chksig%> />&nbsp;By
	checking this box and typing my name below, I am electronically signing
	my application.</p>
</div>
<div id="sig_2">
	<p style="font-weight: bold;">Step 2: Type in your name</p>
	<input style="width: 25%;" type="text" placeholder="First Name" id="fname" name="fname" />&nbsp;
	<input style="width: 25%;" type="text" placeholder="Middle Name" id="mname" name="mname" />&nbsp;
	<input style="width: 25%;" type="text" placeholder="Last Name" id="lname" name="lname" />&nbsp;
	<input style="width: 15%;" type="text" placeholder="Suffix" id="suffix" name="suffix" />
</div>
<div id="sig_3" style="text-align: center; margin: 15px 0px;">
	<button class="button button-primary" style="width: 50%;" id="btnOK" name="btnOK">Submit</button>
</div>
			</div>
			<div id="sigplace">
				(signed electronically)<br />
				on <strong><%= strLast %></strong>
				from [<strong><%= strIP %></strong>]
				<br />
				user agent: <strong><%=strUA%></strong>
			</div>
		</div>
	</div>
		</div>
	</div>

	<div class="row" style="margin-top: 50px;">
		<div class="one column">&nbsp;</div>
		<div class="ten columns" style="border-top: 1px dotted #999;">
			&nbsp;
		</div>
	</div>
	</form>
	</div>
</body>
</html>