<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
Set rsAv = Server.CreateObject("ADODB.RecordSet") 
	ctr = 0
	Do Until ctr = Cint(Request("hctr"))
		sqlAv = "SELECT accept, ansTS FROM AppT_T WHERE appID = " & Request("hid" & ctr) & " AND IntrID = " & Session("UIntr") & " AND accept = 0"
		rsAv.Open sqlAv, g_strCONN, 1, 3
		If Not rsAv.EOF Then
			rsAv("accept") = Request("selSagot" & ctr)
			If Request("selSagot" & ctr) = 1 Or Request("selSagot" & ctr) = 2 Then rsAv("ansTS") = Now
			rsAv.Update
		End If
		rsAv.Close
		ctr = ctr + 1
	Loop
	Set rsAv = Nothing
	Session("MSG") = "Availability Saved."
	Response.Redirect "availappt.asp"
%>