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
	", COALESCE(r.[signature], '') AS [signature]" & _
	", i.[First Name] + ' ' + i.[Last Name] AS [inter_name]" & _
	"FROM [survey2018]				AS y " & _
	"INNER JOIN [interpreter_T]		AS i ON y.[iid]=i.[index] " & _
	"INNER JOIN [user_T]			AS u ON y.[uid]=u.[index] " & _
	"LEFT JOIN  [survey2018med]		AS m ON y.[iid]=m.[iid] " & _
	"LEFT JOIN  [surveyreports] 	AS r ON y.[iid]=r.[iid] " & _	
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
	dtSig = Z_MDYDate( rsSurv("signature") )
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

If (Z_CDate(dtSig) < CDate("2018-01-01")) Then dtSig = "_______________"

If lngIdx <= 0 Then
	Session("MSG") = "not enough survey resources to create a report -- must have at least one!"
	Response.Redirect "survey.list.asp"
End If

avgPunct = avgPunct / lngIdx
avgProfb = avgProfb / lngIdx
avgProcG = avgProcG / lngIdx
avgTeamW = avgTeamW / lngIdx
avgProDv = avgProDv / lngIdx
avgOvral = (avgPunct + avgProfb + avgProcG + avgTeamW + avgProDv) / 5

styPunct = Int(avgPunct)
styProfb = Int(avgProfb)
styProcG = Int(avgProcG)
styTeamW = Int(avgTeamW)
styProDv = Int(avgProDv)

%>
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
					<tr><td>Completed the required trainings in Relias:  <div class="resp"><%=avgReliasTrng%></div>
						</td><td></td></tr>
					<tr><td><h5>Overall Rating</h5></td>
						<td class="resp" style="border: 1px solid #888 !important;"><%=avgOvral%></td>
					</tr>
				</tbody>
			</table>
			<div style="page-break-before:always"></div>
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