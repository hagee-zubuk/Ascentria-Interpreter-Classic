<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
uploadPathtest = "\\10.10.16.35\Interpreter_Upload\"
Function MakeNewFileName()
	strNow = Now
	strNow = Replace(strNow, "/", "")
	strNow = Replace(strNow, ":", "")
	strNow = Replace(strNow, " ", "")
	MakeNewFileName = strNow
End Function
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	Server.ScriptTimeout = 10800
	'upload file to server
	Set oUpload = Server.CreateObject("SCUpload.Upload")
	oUpload.Upload
	If oUpload.Files.Count = 0 Then
		Set oUpload = Nothing
		Session("MSG") = "Please specify a file to import."
		Response.Redirect "uploadtest.asp"
	End If
	oFileName = oUpload.Files(1).Item(1).filename
	If Z_GetExt(oFileName) <> "CSV" Then
		Set oUpload = Nothing
		Session("MSG") = "Invalid File."
		Response.Redirect "uploadtest.asp"
	End If
	nFileName = MakeNewFileName() & ".CSV"
response.write "path: " & uploadpathtest & " filename: " & nfilename
	oUpload.Files(1).Item(1).Save uploadPathtest, nFileName
	Set oUpload = Nothing
response.write "SUCCESS"
End If
%>
<html>
	<head>
		<title>upload test</title>
		<script language='JavaScript'>
			
		</script>
		
	</head>
	<body bgcolor='white' LEFTMARGIN='0' TOPMARGIN='0'>
		<form method="POST" enctype="multipart/form-data">
		
			<table border='0' align='center'>
				<tr><td colspan='2' align='center'><font color='red'  face='trebuchet MS' size='1'>&nbsp;<%=Session("MSG")%>&nbsp;</font></td></tr>
				<tr>
					<td align='center'>
						<input type="file" name="F1" size="20">
					</td>
				</tr>
				<td align='center'>
						<input type='submit' value='upload' style='width: 200px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
					</td>
				</tr>
			</table>
			
		</form>
	</body>
</html>
<%
Session("MSG") = "" 
%>