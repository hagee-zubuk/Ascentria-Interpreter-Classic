<!DOCTYPE html>
<html lang="en">
<%language=vbscript%>
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
			"FROM [msd_form] WHERE [userid]=" & fetchid
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
h3 { font-size: 1.3em; font-weight: bold; }
div.refns { margin: 5px 20px 20px; font-size: 90%; }
div.inputneeded { background-color: lemonchiffon; padding: 15px 20px 5px; margin: 0px; }
.inputneeded p { margin: 0px; }
div.row { margin-top: 20px; }
ol.letters { list-style-type: lower-alpha; counter-reset: list; }
ol.letters > li { list-style: none; }
ol.letters > li:before {
    counter-increment: list;
    content: "(" counter(list, lower-alpha) ") ";
</style>
</head>
<body>
	<div class="container">
	<div class="row" style="text-align: center;">
		<img src='images/LBISLOGO.jpg' border="0" style="width: 287px; height: 64px;" />
		<div class="twelve columns"><h1>Manchester School District Confidentiality and<br />
		Data Security Agreements for Language Bank Interpreters</h1></div>
	</div>
<form id="frmMain" action="msd_ia_proc.asp" method="post">
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="ten columns" style="border-top: 1px dotted #999;">
			&nbsp;
			<input type="hidden" name="empname"	id="empname"	value="<%=strEmpName%>"	/>
			<input type="hidden" name="addr"	id="addr"		value="<%=strAddr%>"	/>
			<input type="hidden" name="userid"	id="userid"		value="<%=lngUserID%>"	/>
			<input type="hidden" name="cellno"	id="cellno"		value="<%=strCellno%>"	/>
			<input type="hidden" name="email"	id="email"		value="<%=strEmail%>"	/>
		</div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="ten columns">
<h2>Personnel 104 CONFIDENTIALITY</h2>
<p>It is the policy of the Manchester School District to respect the privacy, dignity, and
confidentiality of all students attending the Manchester School District. This policy covers
student records, medical information, and other personally identifiable sources of information. It
is the policy of the Manchester School District that such personally identifiable information
should only be viewed or received by School District employees who have a legitimate
educational interest in viewing or receiving the information, as well as those officials involved in
a supervisory capacity over the school in which the students are enrolled. All Manchester
School District employees shall comply with the Family Educational Rights and Privacy Act
(FERPA).</p>
The Superintendent shall develop regulations for the implementation of this policy.<br />
Legal Reference: Family Educational Rights and Privacy Act (FERPA)
<div class="refns">
	First Reading Coordination Committee: 5/11/10<br />
	Second Reading Coordination Committee: 8/10/10<br />
	Third Reading and Adoption by BOSC: 8/23/10
</div>
<div class="inputneeded" id="check1">
<p><input type="checkbox" class="checkme" id="chk_01" name="chk_01" value="1" />
 Check here to acknowledge that you’ve read and understand the above statements
</p>
</div>
		</div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="ten columns">
<h2>Safety 124 DATA MANAGEMENT<br />
Also Safety 123<br />
(Public Use of School Records)</h2>
<p>The Superintendent and/or his/her designee is hereby designated the custodian of all records,
minutes, documents, writings, letters, memoranda, or other written, typed, copied, or developed
materials possessed, assembled, or maintained by this District.</p>
<ol>
<li>All requests for public information are to be forwarded to the Superintendent and/or
his/her designee immediately upon receipt. The Superintendent and/or his/her designee
shall thereupon make a determination as to whether or not the information requested is
public in nature, always considering the privacy requirements of FERPA and IDEA. If
public, the Superintendent and/or his/her designee shall provide the information in a
timely manner which does not disrupt the operation of the schools</li>
<li>If the Superintendent and/or his/her designee finds the information is not confidential
pursuant to FERPA or IDEA and if the information is public in nature pursuant to RSA
91-A:4, he or she shall direct that it be reproduced on the premises. The party
requesting the information shall be charged the cost of reproduction and any other
expenses entailed in locating and retrieving the information. If the information is in active
use or otherwise unavailable, the party requesting the information shall be notified
immediately upon it becoming available.</li>
<li>If the Superintendent and/or his/her designee finds the information to not be public in
nature, he or she shall so inform the requesting party and shall for no reason release
such information.</li>
<li>If the Superintendent and/or his/her designee is unable to ascertain whether or not the
information requested is public in nature, he or she is hereby authorized to request, on
behalf of the Board, an opinion from the Board's attorney as to the nature of the
information. Such opinion requests shall be made within ten (10) days of the original
request for the information. The Superintendent and/or his/her designee shall notify the
person requesting such information that an opinion is being requested of the attorney
and shall notify such person immediately upon receipt of an answer from the attorney.</li>
<li>Student confidentiality shall be maintained at all times in the release of any information
that is determined to be public in nature.</li>
</ol>
Statutory Reference:<br />
<div class="refns">
RSA 91-A:4 (Minutes and Records Available for Public Inspection<br />
NHSBA Code EH<br />
08/94 revised as follows<br />
First Reading Coordination: 02/13/02<br />
Second Reading and Approval BOSC: 03/11/02<br />
</div>
<div class="inputneeded"  id="check2">
<p><input type="checkbox" class="checkme" id="chk_02" name="chk_02" value="1" />
 Check this box to acknowledge that you’ve read and understand the above statements.
</p></div>
		</div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="ten columns">
<h2>Safety 134: STANDARDS FOR THE PROTECTION OF PERSONAL INFORMATION OF<br />
STAFF AND STUDENTS</h2>	
<p>The Superintendent and/or his/her designee shall develop, implement, and maintain a
comprehensive information security program that is written in one or more readily accessible
parts and contains administrative, technical, and physical safeguards that are appropriate to (a)
the size, scope and type of entity of the Manchester School District; (b) the amount of resources
available to the Manchester School District; (c) the amount of stored data; and (d) the need for
security and confidentiality of both student and employee information.</p>

<h3>Scope:</h3>
<p>The provisions of this policy apply to all persons that own, license or have access to personal
information about an employee or student of the Manchester School District.</p>
<ol>
<li>The comprehensive information security program will establish standards to be met in
connection with the safeguarding of personal information contained in both paper and electronic
records. The objectives of this comprehensive security program are to insure the security and
confidentiality of personal information in a manner fully consistent with industry standards;
protect against anticipated threats or hazards to the security or integrity of such information; and
protect against unauthorized access to or use of such information that may result in substantial
harm or inconvenience to any employee or student.</li>
<li>Without limiting the generality of the foregoing, the comprehensive information security
program shall include, but shall not be limited to:
	<ol class="letters">
	<li>Designating one or more employees to maintain the comprehensive
	information security program;</li>
	<li>Identifying and assessing reasonably foreseeable internal and external risks to the
security, confidentiality, and/or integrity of any electronic, paper or other records
containing personal information, and evaluating and improving, where necessary, the
effectiveness of the current safeguards for limiting such risks, including but not limited to:
		<ol>
		<li>ongoing employee (including temporary and contract employee) training;</li>
		<li>employee compliance with policies and procedures; and</li>
		<li>means for detecting and preventing security system failures.</li>
		</ol>
	</li>
	<li>Developing security regulations for employees relating to the storage, access
		and transportation of records containing personal information outside of school
		premises.</li>
	<li>Imposing disciplinary measures, up to and including termination, for violations
		of the comprehensive information security program rules and regulations.</li>
	<li>Preventing terminated employees from accessing records containing personal
		information.</li>
	<li>Oversee service providers, by:
		<ol>
		<li>Taking reasonable steps to select and retain third-party service 
			providers that are capable of maintaining appropriate security measures
			to protect such personal information consistent with this policy and any
			applicable federal policies; and</li>
		<li>Requiring such third-party service providers by contract to implement and
			maintain such appropriate security measures for personal information.</li>
		</ol></li>
	<li>Reasonable restrictions upon physical access to records containing personal
		information, and storage of such records and data in locked facilities,
		storage areas or containers.</li>
	<li>Regular monitoring to ensure that the comprehensive information security
		program is operating in a manner reasonably calculated to prevent unauthorized
		access to or unauthorized use of personal information; and upgrading information
		safeguards as necessary to limit risks.</li>
	<li>Reviewing the scope of the security measures at least annually or whenever there
		is a material change in business practices that may reasonably implicate the
		security or integrity of records containing personal information.</li>
	<li>Documenting responsive actions taken in connection with any incident involving
		a breach of security, and mandatory post-incident review of events and actions
		taken, if any, to make changes in business practices relating to protection of
		personal information.</li>
	</ol>
</li></ol>
<h2>Computer System Security Requirements</h2>
<p>The Superintendent and/or his/her designee shall include in its comprehensive information
security program the establishment and maintenance of a security system covering its
computers, including any wireless system, that, at a minimum, and to the extent technically
feasible, shall have the following elements:</p>
	<ol>
	<li>Secure user authentication protocols including:
		<ol class="letters">
		<li>control of user IDs and other identifiers;</li>
		<li>a reasonably secure method of assigning and selecting passwords,</li>
		<li>control of data security passwords to ensure that such passwords
			are kept in a location and/or format that does not compromise the
			security of the data they protect;</li>
		<li>restricting access to active users and active user accounts only;
			and</li>
		<li>blocking access to user identification after multiple unsuccessful
			attempts to gain access or the limitation placed on access for the
			particular system;</li>
		</ol></li>
	<li>Secure access control measures that:
		<ol class="letters">
		<li>restrict access to records and files containing personal information
			to those who need such information to perform their job duties;
			and</li>
		<li>assign unique identifications plus passwords, which are not vendor
			supplied default passwords, to each person with computer access, that
			are reasonably designed to maintain the integrity of the security of
			the access controls;</li>
		<li>Encryption of all transmitted records and files containing personal
			information that will travel across public networks, and encryption
			of all data containing personal information to be transmitted
			wirelessly.</li>
		<li>Reasonable monitoring of systems, for unauthorized use of or access 
			to personal information;</li>
		<li>Encryption of all personal information stored on laptops or other
			portable devices;</li>
		<li>For files containing personal information on a system that is
			connected to the Internet, there must be reasonably up-to-date
			firewall protection and operating system security patches,
			reasonably designed to maintain the integrity of the personal
			information.</li>
		<li>Reasonably up-to-date versions of system security agent software
			which must include malware protection and reasonably up-to-date
			patches and virus definitions, or a version of such software that
			can still be supported with up-to-date patches and virus definitions,
			and is set to receive the most current security updates on a 
			regular basis.</li>
		<li>Education and training of employees on the proper use of the computer
			security system and the importance of personal information security.</li>
		</ol></li>
	</ol>
<h3>References:</h3>
<p>
This Data Privacy Policy has been written with the assistance of the following:</p>
<div class="refns">
	RSA 189:67 Limits on Disclosure of Information.<br />
	Family Educational Rights and Privacy Act (FERPA):<br />
	http://www.ed.gov/policy/gen/guid/fpco/ferpa/index.html<br />
	CoSN Protecting Privacy in Connected Learning Toolkit<br />
	http://www.cosn.org/focus-areas/leadership-vision/protecting-privacy<br />
	201 CMR 17.00: STANDARDS FOR THE PROTECTION OF PERSONAL
	INFORMATION OF RESIDENTS OF THE COMMONWEALTH OF MASSACHUSETTS<br />
	Privacy Pitfalls as Education Apps Spread Haphazardly<br />
	http://www.nytimes.com/2015/03/12/technology/learning-apps-outstrip-school-oversightand-student-privacy-is-among-the-risks.html?_r=1<br />
	RSA 189:67 Limits on Disclosure of Information.<br />
	First Reading, IT Committee: 11/23/15<br />
	Second Reading, Coordination: 12/14/15<br />
	Adoption, BOSC: 1/25/16
</div>
<h2>Student and Teacher Information Protection and Privacy</h2>
<h3>Section 189:67</h3>	
<p>189:67 Limits on Disclosure of Information. –</p>
	<ol type="I">
	<li>A school shall, on request, disclose student personally-identifiable
		data about a student to the parent, foster parent, or legal guardian
		of the student under the age of 18 or to the eligible student.</li>
	<li>A school or the department may disclose to a testing entity the
		student's name or unique pupil identifier, but not both, and birth 
		date for the sole purpose of identifying the test taker. This data
		shall be destroyed by the testing entity as soon as the testing
		entity has completed the verification of test takers, shall not be
		disclosed by the testing entity to any other person, organization,
		entity or government or any component thereof, other than the
		district, school or school district, and shall not be used by the
		testing entity for any other purpose whatsoever, including but not
		limited to test-data analysis.</li>
	<li>Neither a school nor the department shall disclose or permit the
		disclosure of student or teacher personally-identifiable data, the
		unique pupil identifier, or any other data listed in RSA 189:68, I
		to any testing entity performing testdata analysis. The testing
		entity may perform the test analysis but shall not connect such 
		data to other student data.</li>
	<li>Except as provided in RSA 193-E:5, or pursuant to a court order
		signed by a judge, the department shall not disclose student
		personally-identifiable data in the SLDS or teacher personally-
		identifiable data in other department data systems to any
		individual, person, organization, entity, government or component
		thereof, but may disclose such data to the school district in which
		the student resides or the teacher is employed.</li>
	<li>Student personally-identifiable data shall be considered confidential
		and privileged and shall not be disclosed, directly or indirectly,
		as a result of administrative or judicial proceedings.</li>
	<li>The department shall report quarterly on its website the number of
		times it disclosed student personally-identifiable data to any person,
		organization entity or government or a component thereof, other than
		the student, his or her parents, foster parents or legal guardian and
		the school district, early childhood program or post-secondary
		institution in which the student wasenrolled at the time of
		disclosure; the name of the recipient or entity of the disclosure; and
		the legal basis for the disclosure.</li>
	</ol>
<p>Source. 2014, 68:1, eff. July 1, 2014. 2015, 71:3, eff. Aug. 1, 2015.</p>

<div class="inputneeded">
	<div id="check3">
<p><input type="checkbox" class="checkme" id="chk_03" name="chk_03" value="1" />
 Check the box to acknowledge that you’ve read and understand the above statements.
</p>
	</div>

<div id="sig_1">
	<p style="font-weight: bold;">Check the box below</p>
	<p style="text-align: center;"><input type="checkbox" value="1" id="chkSig" name="chkSig" <%=chksig%> />&nbsp;By
	checking this box and typing my name below, I am electronically signing this document
	my application.</p>
</div>
<div id="sig_2">
	<p style="font-weight: bold;">Type in your name</p>
	<input style="width: 25%;" type="text" placeholder="First Name" id="fname" name="fname" />&nbsp;
	<input style="width: 25%;" type="text" placeholder="Middle Name" id="mname" name="mname" />&nbsp;
	<input style="width: 25%;" type="text" placeholder="Last Name" id="lname" name="lname" />&nbsp;
	<input style="width: 15%;" type="text" placeholder="Suffix" id="suffix" name="suffix" />
</div>
<div id="sig_3" style="text-align: center; margin: 15px 0px;">
	<button type="button" class="button button-primary" style="width: 50%;" id="btnOK" name="btnOK">Submit</button>
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
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="ten columns">
		</div>
	</div>
</form>
	</div> <!-- container -->
</body>
</html>
<script><!--
$(document).ready(function(){
		$('#sigplace').hide();
		$('#sigsection').show();
		$('#sig_2').hide();
		$('#sig_3').hide();
		$('#chkSig').attr('checked', false);
		$('#chkSig').change(function(){
				if ($('#chkSig').prop('checked')) {
					$('#sig_2').show();
					$('#sig_3').show();
					$('#name_first').focus();
				} else {
					$('#sig_2').hide();
					$('#sig_3').hide();
					$('#check1').hide();
					$('#check2').hide();
					$('#check3').hide();
				}
			});
		$('#btnOK').click(function(){
				var blnOK = $('#chkSig').is(":checked") && 
						$('#chk_01').is(":checked") && 
						$('#chk_02').is(":checked") && 
						$('#chk_03').is(":checked")
						;
				
				console.log( "Checkboxes: " + blnOK );
				if (!blnOK) {
					if(! $('#chk_01').is(":checked")) {
						var ofst = $("#chk_01").offset();
						window.scrollTo(0,ofst.top -20);
					}
					if(! $('#chk_02').is(":checked")) {
						var ofst = $("#chk_02").position();
						window.scrollTo(0,ofst.top - 40);
					}
					alert("You have to check all the checkboxes to sign.");
					return false;
				} else {
					var zzs = $.trim( $('#fname').val() + $('#lname').val() );
					if (zzs.length <= 1) {
						alert("Please type in your full name to sign.");
						return false;
					}
					console.log( "OK! We're going!" ); console.log();
					$('#frmMain').submit();
				}

			});
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
	Response.Write "alert(""" & tmpMSG & """);"
	Session("MSG") = ""
End if
%>
	});
--></script>
