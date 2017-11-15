<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
'save datestamp

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
'create vform
Set theDoc = Server.CreateObject("ABCpdf6.Doc") 'converts html to pdf
	attachPDF = pdfStr & "VerificationForm" & Request("ReqID") & ".pdf"
	strUrl = "https://lbis.lssne.org/interpreter/print.asp?PDF=1&ID=" & Request("ReqID")
	'strUrl = "http://webserv2/lss-lbis/print.asp?PDF=1&ID=" & Request("HID")
	'strUrl = "http://web03.zubuk.com/lss-lbis/print.asp?PDF=1&ID=" & Request("HID")
	thedoc.HtmlOptions.PageCacheClear
	theDoc.HtmlOptions.RetryCount = 3
	theDoc.HtmlOptions.Timeout = 120000
	theDoc.Pos.X = 10
	theDoc.Pos.Y = 10
	theID = theDoc.AddImageUrl(strUrl)
	
	Do
	  If Not theDoc.Chainable(theID) Then Exit Do
	  theDoc.Page = theDoc.AddPage()
	  theID = theDoc.AddImageToChain(theID)
	Loop
	
	For i = 1 To theDoc.PageCount
	  theDoc.PageNumber = i
	  theDoc.Flatten
	Next

	theDoc.Save attachPDF

'downloadfile
Set dload = Server.CreateObject("SCUpload.Upload")
			tmpfile = attachPDF'"C:\work\LSS-LBIS\web-DMZ\Images\icon_download.gif" 'attachPDF
			dload.Download tmpFile
			Set dload = Nothing
			download = 1
			rsUser.Close
			Set rsUser = Nothing
End If
%>
<html>
	<head>
		<title>DMZ - DL TEST</title>
	</head>
	<body>
		<form name="frmTest" method="POST" action="dltest.asp">
			<input type="text" name="reqID" /><input type="submit" name="DL">
		</form>
	</body>
</html>