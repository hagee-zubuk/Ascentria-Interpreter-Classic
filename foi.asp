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
	strSQL = "SELECT [userid], [empname], [addr], [cellno], [email], [fname], [mname], [lname], [suffix], [chksig], [ip], [useragent], [last] FROM [InfoRelease] WHERE [userid]=" & fetchid
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
<!doctype html>
<html lang="en">
<%language=vbscript%>
<head>
	<meta charset="utf-8">
	<title>Authorization for Release of Information</title>
	<meta name="description" content="Authorization for Release of Information">
	<meta name="author" content="Argao.net">
<%
If (blnDisp) Then
	'meh
Else
%>
	<script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" crossorigin="anonymous"
			integrity="sha256-k2WSCIexGzOj3Euiig+TlR8gA0EmPjuc79OEeY5L45g="></script>
	<script src="foi.js"></script>
<%
End If
%>
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/normalize/7.0.0/normalize.css"
			integrity="sha256-sxzrkOPuwljiVGWbxViPJ14ZewXLJHFJDn0bv+5hsDY=" crossorigin="anonymous" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/skeleton/2.0.4/skeleton.css"
			integrity="sha256-ECB9bbROLGm8wOoEbHcHRxlHgzGqYpDtNTgDTyDz0wg=" crossorigin="anonymous" />
</head>
<body>
<div class="container"><form id="frmROI" name="frmROI" method="post" action="foi_proc.asp">
	<div class="row" style="text-align: center;">
		<img src='images/LBISLOGO.jpg' border="0" style="width: 287px; height: 64px;" />
		<div class="twelve columns"><h1>RELEASE  OF  INFORMATION</h1></div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="six columns">
			<label for="empName">Name of Employee:</label>
			<input class="u-full-width" type="text" placeholder="Employee Name" id="empname" name="empname"
					value="<%= strEmpName%>"
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
			<input class="u-full-width" type="tel" maxlength="20" placeholder="Cell#" id="cellno" name="cellno"
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
			<p>I authorize Language Bank, a member of Ascentria Care Alliance, to share my personal
			information described below with Language Bank customers upon reasonable request of
			a customer for the purpose of compliance audits, credentialing monitoring, or public
			health concerns.  Depending on the nature of the request, Language Bank may disclose
			the following information responsive to that request:</p>
<ul style="margin-left: 10%;">
	<li>resume</li>
	<li>educational background</li>
	<li>credentialing certifications for interpretation</li>
	<li>continuing education certifications</li>
	<li>performance evaluations and competency check list</li>
	<li>immunization records</li>
	<li>contact information (including mobile phone number and e-mail address)</li>
</ul>
			<p>I understand that Language Bank will make all reasonable efforts to maintain my privacy
			in connection with the disclosure of this information, and that customers will be advised
			not to contact me directly outside of the time of scheduled appointments or outside the
			scope of standard business relations.</p>

			<p>My signature below releases Language Bank and Ascentria Care Alliance and their
			authorized agents from all liability in connection with the disclosure of this information.
			I understand that I may revoke my consent by sending written notice of revocation to the
			Language Bank Program Manager or Assistant Manager.  Such revocation will apply only to
			the release of information regarding appointments which have yet to occur as of the date
			of receipt of the revocation.  My revocation will not affect the release of information
			about appointments that have already occurred.</p>
			<p>Date:<%= FormatDateTime(Now(), 0)%></p>
			<div id="sigsection" style="border: 1px dotted #777; padding: 5px; display: none;">
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
	<div class="row" style="margin-top: 50px;">
		<div class="one column">&nbsp;</div>
		<div class="ten columns" style="border-top: 1px dotted #999;">
			&nbsp;
		</div>
	</div>
</form></div>
</body>
</html>