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
strDest = "survey.results.asp"
If Not rsSurv.EOF Then
	blnRelease = CBool( rsSurv("release") )
	If blnRelease Then
		'Response.Write "RELEASED<br /><br />"
		rsSurv("signature") = Now
		rsSurv.Update
		strDest = "survey.report.asp"
	End If
End If
rsSurv.Close
Set rsSurv = Nothing
If Not blnRelease Then
	Session("MSG") = "survey response index is missing"
	strDest "survey.v18.asp"
End If

Response.Redirect strDest
'Response.Write strDest
%>
