<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
Function ChkBoxState(aaa)
	ChkBoxState = "&mdash;"
	valA = Z_CLng(aaa)
	If valA = 1 Then
		ChkBoxState = "<img src=""images/check.png"" title=""X"" alt=""V"" />"
	End If
End Function

lngID = Session("UIntr")
If lngID < 1 Then
	lngID = CLng(Request("ix"))
	If lngID < 1 Then
		Session("MSG") = "survey response index is missing"
		Response.Redirect "survey.v18.asp"
	End If
End If

Set rsSurv = Server.CreateObject("ADODB.RecordSet")
strSQL = "SELECT [release] FROM [surveyreports] WHERE [iid]=" & lngID
rsSurv.Open strSQL, g_strCONN, 3, 1
blnRelease = FALSE
If Not rsSurv.EOF Then
	blnRelease = CBool( rsSurv("release") )
End If
rsSurv.Close
Set rsSurv = Nothing

If Not blnRelease Then
	Session("MSG") = "survey response index is missing"
	Response.Redirect "survey.v18.asp"
End If

Set rsSurv = Server.CreateObject("ADODB.RecordSet")
strSQL = "SELECT y.[index]" & _
	", y.[rdoPunct], y.[rdoProfb], y.[rdoProcG], y.[rdoTeamW], y.[rdoProDv], y.[rdoReliasTrng]" & _
	", y.[txtGoals], y.[txtStrengths], y.[txtImprovement], y.[txtComments]" & _
	", y.[iid]" & _
	", COALESCE(m.[index], 0) AS [med_ix]" & _
	", i.[First Name] + ' ' + i.[Last Name] AS [inter_name]" & _
	"FROM [survey2018]				AS y " & _
	"INNER JOIN [interpreter_T]		AS i ON y.[iid]=i.[index] " & _
	"INNER JOIN [user_T]			AS u ON y.[uid]=u.[index] " & _
	"LEFT JOIN  [survey2018med]		AS m ON y.[iid]=m.[iid] " & _
	"WHERE y.[iid]=" & lngID
rsSurv.Open strSQL, g_strCONN, 3, 1
If rsSurv.EOF Then
	rsSurv.Close
	Set rsSurv = Nothing
	Session("MSG") = "survey response index was not found"
	Response.Redirect "survey.list.asp"
End If
lngIdx = 0
avgPunct = 0
avgProfb = 0
avgProcG = 0
avgTeamW = 0
avgProDv = 0
txtGoals = ""
txtStrng = ""
txtImprv = ""
txtComnt = ""
avgReliasTrng = "N"
Do While Not rsSurv.EOF
	txtInterpreter = rsSurv("inter_name")
	avgPunct = avgPunct + Z_CLng(rsSurv("rdoPunct"))
	avgProfb = avgProfb + Z_CLng(rsSurv("rdoPunct"))
	avgProcG = avgProcG + Z_CLng(rsSurv("rdoProcG"))
	avgTeamW = avgTeamW + Z_CLng(rsSurv("rdoTeamW"))
	avgProDv = avgProDv + Z_CLng(rsSurv("rdoProDv"))
	If rsSurv("rdoReliasTrng") = "Y" Then avgReliasTrng = "Y"
	If Len(Z_FixNull(rsSurv("txtGoals"))) > 0 Then txtGoals = txtGoals & rsSurv("txtGoals") & vbCrLf
	If Len(Z_FixNull(rsSurv("txtStrengths"))) > 0 Then txtStrng = txtStrng & rsSurv("txtStrengths") & vbCrLf
	If Len(Z_FixNull(rsSurv("txtImprovement"))) > 0 Then txtImprv = txtImprv & rsSurv("txtImprovement") & vbCrLf
	If Len(Z_FixNull(rsSurv("txtComments"))) > 0 Then txtComnt = txtComnt & rsSurv("txtComments") & vbCrLf

	lngMedIx = CLng(rsSurv("med_ix"))
	' iterate!
	rsSurv.MoveNext
	lngIdx = lngIdx + 1
Loop
rsSurv.Close
Set rsSurv = Nothing

If lngIdx <= 0 Then
	Session("MSG") = "not enough survey resources to create a report -- must have at least one!"
	Response.Redirect "survey.list.asp"
End If

avgPunct = avgPunct / lngIdx
avgProfb = avgProfb / lngIdx
avgProcG = avgProcG / lngIdx
avgTeamW = avgTeamW / lngIdx
avgProDv = avgProDv / lngIdx

styPunct = Int(avgPunct)
styProfb = Int(avgProfb)
styProcG = Int(avgProcG)
styTeamW = Int(avgTeamW)
styProDv = Int(avgProDv)

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
	<div class="row" id="intrbar">
		<div class="five columns">
			<b>Interpreter Name</b>&nbsp;&nbsp;<div style="display: inline-block;font-weight: bold; font-size: 150%;"><%=txtInterpreter%></div>
		</div>
		<div class="seven columns align-right no-print">
			<button type="button" class="button button-primary"id="btnPDF" name="btnPDF">PDF</button>
		</div>
	</div>
	<div class="row">
		<div class="twelve columns">
			<table class="u-full-width smallertable">
  				<thead>
    				<tr><th colspan="2" class="yellow">Performance Criteria</th></tr>
  				</thead>
  				<tbody>
  					<tr><td><h5>Punctuality</h5>
							</td>
						<td class="resp rr<%=styPunct%>"><%=avgPunct%></td>
					</tr>
					<tr><td><h5>Professional Behavior</h5>
							</td>
						<td class="resp rr<%=styProfb%>"><%=avgProfb%></td>
					</tr>
					<tr><td><h5>Adherence to LB Procedural Guidelines</h5>
							</td>
						<td class="resp rr<%=styProcG%>"><%=avgProcG%></td>
					</tr>
					<tr><td><h5>Team Work Ethics</h5>
							</td>
						<td class="resp rr<%=styTeamW%>"><%=avgTeamW%></td>
					</tr>
					<tr><td><h5>Professional Development</h5></td>
						<td class="resp rr<%=styProDv%>"><%=avgProDv%></td>
					</tr>
				</tbody>
			</table>
			Completed the required trainings in Relias (Yes or No):  <div class="resp"><%=avgReliasTrng%></div>
			<br /><br />
			<p>
				<label>Goals:</label>
				<pre class="resp"><%=txtGoals%></pre>
			</p>
			<p>
				<label>Existing Strengths:</label>
				<pre class="resp"><%=txtStrng%></pre>
			</p>
			<p>
				<label>Areas Needing Improvement:</label>
				<pre class="resp"><%=txtImprv %></pre>
			</p>
			<p>
				<label>Comments:</label>
				<pre class="resp"><%=txtComnt %></pre>
			</p>
  		</div>
  	</div>
<%
strSQL = "SELECT [index] AS [id], [First Name] + ' ' + [Last Name] AS [intr] " & _
		"FROM [interpreter_T] " & _
		"WHERE [index]= " & lngID
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
rsIntr.Open strSQL, g_strCONN, 3, 1
If Not rsIntr.EOF Then
	txtName = Z_FixNull(rsIntr("intr"))
	lngID = Z_FixNull( rsIntr("id") )
End If
rsIntr.Close
Set rsIntr = Nothing
blnLoad = TRUE
blnNm = " readonly=""readonly"" "
strSQL = "SELECT * FROM [survey2018med] WHERE [iid]=" & lngID 
Set rsSurv = Server.CreateObject("ADODB.RecordSet")
rsSurv.Open strSQL, g_strCONN, 3, 1
If Not rsSurv.EOF Then
	blnA1 = ChkBoxState(rsSurv("chkA1"))
	blnA2 = ChkBoxState(rsSurv("chkA2"))
	blnA3 = ChkBoxState(rsSurv("chkA3"))
	blnA4 = ChkBoxState(rsSurv("chkA4"))
	blnA5 = ChkBoxState(rsSurv("chkA5"))
	blnA6 = ChkBoxState(rsSurv("chkA6"))
	blnB1 = ChkBoxState(rsSurv("chkB1"))
	blnC1 = ChkBoxState(rsSurv("chkC1"))
	blnC2 = ChkBoxState(rsSurv("chkC2"))
	blnC3 = ChkBoxState(rsSurv("chkC3"))
	blnC4 = ChkBoxState(rsSurv("chkC4"))
	blnD1 = ChkBoxState(rsSurv("chkD1"))
	blnD2 = ChkBoxState(rsSurv("chkD2"))
	blnD3 = ChkBoxState(rsSurv("chkD3"))
	blnD4 = ChkBoxState(rsSurv("chkD4"))
	blnD5 = ChkBoxState(rsSurv("chkD5"))
	blnD6 = ChkBoxState(rsSurv("chkD6"))
	blnD7 = ChkBoxState(rsSurv("chkD7"))
	blnD8 = ChkBoxState(rsSurv("chkD8"))
	blnE1 = ChkBoxState(rsSurv("chkE1"))
	blnE2 = ChkBoxState(rsSurv("chkE2"))
	blnE3 = ChkBoxState(rsSurv("chkE3"))
	blnE4 = ChkBoxState(rsSurv("chkE4"))
	blnE5 = ChkBoxState(rsSurv("chkE5"))
	blnF1 = ChkBoxState(rsSurv("chkF1"))
	blnF2 = ChkBoxState(rsSurv("chkF2"))
	blnF3 = ChkBoxState(rsSurv("chkF3"))
	blnF4 = ChkBoxState(rsSurv("chkF4"))
	blnF5 = ChkBoxState(rsSurv("chkF5"))
	blnF6 = ChkBoxState(rsSurv("chkF6"))
	blnF7 = ChkBoxState(rsSurv("chkF7"))
	blnG1 = ChkBoxState(rsSurv("chkG1"))
	blnG2 = ChkBoxState(rsSurv("chkG2"))
	blnH1 = ChkBoxState(rsSurv("chkH1"))
	blnH2 = ChkBoxState(rsSurv("chkH2"))
	blnH3 = ChkBoxState(rsSurv("chkH3"))
	blnH4 = ChkBoxState(rsSurv("chkH4"))
	blnH5 = ChkBoxState(rsSurv("chkH5"))
	blnH6 = ChkBoxState(rsSurv("chkH6"))
	blnH7 = ChkBoxState(rsSurv("chkH7"))
	blnH8 = ChkBoxState(rsSurv("chkH8"))
	blnH9 = ChkBoxState(rsSurv("chkH9"))
	blnH10 = ChkBoxState(rsSurv("chkH10"))
	blnH11 = ChkBoxState(rsSurv("chkH11"))
	blnH12 = ChkBoxState(rsSurv("chkH12"))
	blnH13 = ChkBoxState(rsSurv("chkH13"))
%>
	<p style="page-break-before: always;">&nbsp;</p>
	<div class="row">
		<div class="twelve columns logobar" style="border-bottom: 1px dashed #bcbcbc;">
			<h1>Medical Interpreter Competency Checklist</h1>
		</div>
	</div>
	<div class="row">
		<div class="ten columns" style="font-size: 9pt;">
			COMPETENCY<br />
			<p>(For further details on competency requirements, please refer
			to Manual of Orientation for Medical Interpreters and Guidelines
			for Establishing Competency)</p>
		</div>
		<div class="two columns" style="font-size: 9pt; text-align: center;vertical-align: bottom;">
			Check if feedback is required			
		</div>
	</div>
	<div class="row">
		<div class="twelve columns">
			<table class="u-full-width smallertable">
  				<thead></thead>
  				<tbody>
    				<tr><th>A. INTRODUCTION/ROLE OF INTERPRETER:  The interpreter...</th><th>&nbsp;</th>
				    </tr>
  					<tr><td class="indent-1">
						Introduces self, explains role of interpreter to patient, and establishes rapport with patient.</td>
						<td><%=blnA1%></td>
					</tr>
					<tr><td class="indent-1">
						Ascertains whether the patient has prior experience working with interpreters.</td>
						<td><%=blnA2%></td>
					</tr>
					<tr><td class="indent-1">
						Encourages patient to ask for clarification of any issue as it arises during the visit.</td>
						<td><%=blnA3%></td>
					</tr>
					<tr><td class="indent-1">
						Relays to the patient legal requirements and essential information regarding informed consent, confidentiality, and security of medical communication.</td>
						<td><%=blnA4%></td>
					</tr>
					<tr><td class="indent-1">
						Asks the provider to introduce him/herself to the patient using his/her full title and to state the provider’s goal for the visit.</td>
						<td><%=blnA5%></td>
					</tr>
					<tr><td class="indent-1">
						Relays to both the health professional and the patient that if either desires a confidential conversation that they do not want the interpreter to hear, that the interpreter must leave the room given the requirement that interpreters translate everything that is said by either the patient or healthcare professional.</td>
						<td><%=blnA6%></td>
					</tr>
					<tr><th>B. MANAGEMENT OF PHYSICAL SPACE: The interpreter...</th><th>&nbsp;</th>
					</tr>
					<tr><td class="indent-1">
						Effectively arranges the spatial configuration of the interview to encourage direct face-to-face contact by the patient and provider of care.</td>
						<td><%=blnB1%></td>
					</tr>
					<tr><th>C. CULTURAL UNDERSTANDING: The interpreter...</th><th>&nbsp;</th>
					</tr>
  					<tr><td class="indent-1">
  						Understands the rules of cultural etiquette with respect to status, age, gender, hierarchy, and level of acculturation.</td>
						<td><%=blnC1%></td>
					</tr>
					<tr><td class="indent-1">
						Demonstrates an understanding of potential barriers to communication including cultural differences, ethnic issues, gender issues, lack of education or differences between patient or provider life experience.</td>
						<td><%=blnC2%></td>
					</tr>
					<tr><td class="indent-1">
						Anticipates the need for and reassesses patient and provider comfort levels and addresses any perceived barriers that may impact on the success of the interaction between provider and patient.</td>
						<td><%=blnC3%></td>
					</tr>
					<tr><td class="indent-1">
						Shares any relevant cultural information with both patient and provider to facilitate understanding between all parties.</td>
						<td><%=blnC4%></td>
					</tr>
					<tr><th>D. INTERPRETATION SKILLS: The interpreter...</th><th>&nbsp;</th>
					</tr>
					<tr><td class="indent-1">Understands the vital role of accurate interpretation and understands the risks of inaccurate interpretation in a medical situation.
					</tr>
					<tr><td class="indent-1">
						Considers and selects the most effective mode of interpretation prior to the start of the interpretation service (e.g., consecutive, simultaneous, or first/third person) and adjusts mode as needed during clinical interview.</td>
						<td><%=blnD1%></td>
					</tr>
					<tr><td class="indent-1">
						Ensures that he/she understands the message prior to transmission.</td>
						<td><%=blnD2%></td>
					</tr>
					<tr><td class="indent-1">
						Understands his/her limitations of medical knowledge, refrains from making assumptions, and demonstrates willingness to obtain clarification of medical terms and concepts as necessary.</td>
						<td><%=blnD3%></td>
					</tr>
					<tr><td class="indent-1">
						Accurately transmits information between patient and provider, transmitting the message completely, utilizing communication aids (e.g., pictures, drawings, or gestures) to supplement communication</td>
						<td><%=blnD4%></td>
					</tr>
					<tr><td class="indent-1">
						Ensures that the listener (patient/family) understands what is being conveyed after transmission of the information.</td>
						<td><%=blnD5%></td>
					</tr>
					<tr><td class="indent-1" colspan="2">
						Manages the flow of communication in order to insure accuracy of transmission and enhance rapport between patient and provider.  Specifically:</td>
					</tr>
					<tr><td class="indent-2">
						Manages the conversation so that only one person talks at a time.</td>
						<td><%=blnD6%></td>
					</tr>
					<tr><td class="indent-2">
						Interrupts the other speaker to allow the other party to speak when necessary.</td>
						<td><%=blnD7%></td>
					</tr>
					<tr><td class="indent-2">
						Indicates clearly when he/she is speaking on his/her own behalf.</td>
						<td><%=blnD8%></td>
					</tr>
					<tr><th>E. COMMUNICATION SKILLS: The interpreter...</th><th>&nbsp;</th></tr>
					<tr><td class="indent-1">
						Is cognizant of the changing tone and emotional content of medical conversations, and remains alert to internal conflicts that may emerge between provider and patient.</td>
						<td><%=blnE1%></td>
					</tr>
					<tr><td class="indent-1">
						When strong feelings or conflict arise between the provider and the patient, the interpreter does not take sides in the conflict and remains calm while acknowledging the tension between patient and provider. He/she manages the situation effectively through use of clarification.</td>
						<td><%=blnE2%></td>
					</tr>
					<tr><td class="indent-1">
						Manages his/her own internal personal conflicts by clearly separating his/her own values and beliefs from those of the patient and provider of care.</td>
						<td><%=blnE3%></td>
					</tr>
					<tr><td class="indent-1">
						Is able to acknowledge openly to the patient/provider that the topic is difficult for interpreter.</td>
						<td><%=blnE4%></td>
					</tr>
					<tr><td class="indent-1">
						Actively identifies his/her own mistakes, corrects him/herself as quickly as possible, communicates that to both patient/provider, and accepts the feedback and restates new understanding for the record.</td>
						<td><%=blnE5%></td>
					</tr>
					<tr><th>F. ROLE AS FACILITATOR: The interpreter...</th><th>&nbsp;</th></tr>
					<tr><td class="indent-1">
						Encourages the provider to give the patient appropriate instructions and makes certain that the patient understands both the instructions and what he/she must do next.</td>
						<td><%=blnF1%></td>
					</tr>
					<tr><td class="indent-1">
						Ascertains from the patient whether he/she has any final questions for the provider.</td>
						<td><%=blnF2%></td>
					</tr>
					<tr><td class="indent-1">
						Assesses whether the patient will need interpretation services after the medical visit is concluded.</td>
						<td><%=blnF3%></td>
					</tr>
					<tr><td class="indent-1">
						Ensures that the patient understands to contact the Provider of Record or OnCall provider, or organization telephone service after hours with any concerns or questions.</td>
						<td><%=blnF4%></td>
					</tr>
					<tr><td class="indent-1">
						Explains after hours process to patients with limited English proficiency.</td>
						<td><%=blnF5%></td>
					</tr>
					<tr><td class="indent-1">
						Ensures appropriate referrals are made, including place, date and time, and ensures interpretive services are scheduled.</td>
						<td><%=blnF6%></td>
					</tr>
					<tr><td class="indent-1">
						Ensures that any concerns raised (before or after the interview) are addressed and referred to clinical personnel who can assist with resolution of such concerns.</td>
						<td><%=blnF7%></td>
					</tr>
					<tr><th>G. ADMINISTRATIVE TASKS: The interpreter...</th><th>&nbsp;</th></tr>
					<tr><td class="indent-1">
						Completes appropriate documentation as indicated or requested by clinical personnel.</td>
						<td><%=blnG1%></td>
					</tr>
					<tr><td class="indent-1">
						Appropriately signs, dates, and indicates time of day on all notes.</td>
						<td><%=blnG2%></td>
					</tr>
					<tr><th>H. ETHICAL STANDARDS: In each of the following areas, the interpreter...</th><th>&nbsp;</th></tr>
					<tr><td class="indent-1" colspan="2">CONFIDENTIALITY:</td></tr>
					<tr><td class="indent-2">
						Is aware of and observes all relevant organizational policies and state/federal laws regarding release of confidential medical information.</td>
						<td><%=blnH1%></td>
					</tr>
					<tr><td class="indent-2">
						Understands that protection of patient confidentiality is NOT limited to the potential for sharing personal medical information outside of the organization, but also includes a prohibition against sharing any of the patient's personal information with anyone on the health care team or in the healthcare organization who does not have a specific need to know that information.</td>
						<td><%=blnH2%></td>
					</tr>
					<tr><td class="indent-1" colspan="2">IMPARTIALITY:</td></tr>
					<tr><td class="indent-2">
						Is aware and able to identify any personal bias, belief, or conflict of interest that may interfere with his/her ability to impartially interpret in any given situation, and discloses this to the provider so that another interpreter can step in to provide the service.</td>
						<td><%=blnH3%></td>
					</tr>
					<tr><td class="indent-1" colspan="2">PROFESSIONAL INTEGRITY:</td></tr>
					<tr><td class="indent-2">
						Acts as a conduit of information, not as an information source, unless specifically trained or licensed to supply that particular information. Therefore, the interpreter refrains from counseling or advising the patient at any time.</td>
						<td><%=blnH4%></td>
					</tr>
					<tr><td class="indent-2">
						Refrains from any contact with the patient outside of employment, avoiding personal benefit</td>
						<td><%=blnH5%></td>
					</tr>
					<tr><td class="indent-2">
						Engages in ongoing professional development.</td>
						<td><%=blnH6%></td>
					</tr>
					<tr><td class="indent-2">
						Maintains professional dress and demeanor at all times.</td>
						<td><%=blnH7%></td>
					</tr>
					<tr><td class="indent-2">
						Is consistently observed to be free of prejudice or critical comments or judgement of the patient.</td>
						<td><%=blnH8%></td>
					</tr>
					<tr><td class="indent-1" colspan="2">PROFESSIONAL DISTANCE:</td></tr>
					<tr><td class="indent-2">
						Can explain the meaning of “distance” in this context, and its implications and consequences.</td>
						<td><%=blnH9%></td>
					</tr>
					<tr><td class="indent-2">
						Refrains from becoming personally involved in the patient's life.</td>
						<td><%=blnH10%></td>
					</tr>
					<tr><td class="indent-2">
						Does not create any expectations that the interpreter role cannot fulfill.</td>
						<td><%=blnH11%></td>
					</tr>
					<tr><td class="indent-2">
						Actively promotes patient self-sufficiency.</td>
						<td><%=blnH12%></td>
					</tr>
					<tr><td class="indent-2">
						Monitors own personal agenda of service, and is aware of transference and countertransference issues, discussing them with the team leader or with his/her supervisor when any boundary issue or potential overreaching of mission could occur.</td>
						<td><%=blnH13%></td>
					</tr>
				</tbody>
			</table>
		</div>

<%	
End If
rsSurv.Close
Set rsSurv = Nothing
%>
<!-- we're done -->
</div>
</body>
</html>
<script language="javascript" type="text/javascript"><!--
$( document ).ready(function() {
	$('#btnPDF').click(function(){
		document.location="survey.pdf.asp";
	});
	console.log( "ready!" );
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
	alert("<%=tmpMSG%>");
<%
End If
%>
});
// --></script>