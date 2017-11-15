<html>
<body>
<%
	Set oFileStream = Server.CreateObject("ADODB.Stream")
	oFileStream.Open
	oFileStream.Type = 1 'Binary
	If Request("file") = 1 Then
		tmpfile = "FWA-GCT_FINAL.pdf" '"images/FWA-GCT_FINAL.pdf"
	ElseIf Request("file") = 2 Then
		tmpfile = "Interpreter Guidelines 6-24-2015 10-22 - FINAL APPROVED.pdf"
	End If
	fpath = "c:\work\LSS-LBIS\web\Images\" & tmpfile
	oFileStream.LoadFromFile fpath
	Response.ContentType = "application/pdf"
	Response.AddHeader "Content-Disposition", "inline; filename=" & fpath
	Response.BinaryWrite oFileStream.Read
	oFileStream.Close
	Set oFileStream= Nothing
%>
</body>
</html>