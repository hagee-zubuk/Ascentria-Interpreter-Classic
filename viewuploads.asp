<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
strRID = Z_CLng(Request("reqid"))
lngUID = Z_CLng(Request("uid"))
tmDate = Request.Cookies("tmpDate")
' blah
Set rsTS = Server.CreateObject("ADODB.RecordSet")
sqlTS = "SELECT req.[index], xrs.[statusname], Cfname, Clname, req.[AStarttime], req.[AEndtime]" & _
		", ins.[Facility], dep.[dept], req.[appDate] " & _
		"FROM [request_T] AS req " & _
		"INNER JOIN [xrStatus] AS xrs ON req.[Status]=xrs.[index] " & _
		"INNER JOIN [dept_T] AS dep ON req.[DeptID]=dep.[index] " & _
		"INNER JOIN [institution_T] AS ins ON req.[InstID]=ins.[Index] " & _
		"WHERE req.[index]=" & strRID
rsTS.Open sqlTS, g_strCONN, 3, 1
If rsTS.EOF Then
	rsTS.Close
	Set rsTS = Nothing
	Session("MSG") = "Appointment not found: " & strRID
	Response.Redirect "tsheet.asp?tmpDate=" & tmDate
	Response.Flush
	Response.End
End If

Set rsUploads = Server.CreateObject("ADODB.RecordSet")
lngVFCnt = 0 
lngTlCnt = 0 
strVForm = ""
strTolls = ""
viewpath = ""
strSQL = "SELECT COUNT([uid]) AS [cnt] FROM uploads WHERE RID=" & strRID & " AND [staff]=0"
rsUploads.Open strSQL, g_strCONNupload, 3, 1
If rsUploads.EOF Then
	rsUploads.Close
	Set rsUploads = Nothing
	Session("MSG") = "Upload database cannot be accessed: " & strRID
	Response.Redirect "tsheet.asp?tmpDate=" & tmDate
	Response.Flush
	Response.End
End If
lngCnt = Z_CLng(rsUploads("cnt"))
rsUploads.Close
Set rsUploads = Nothing
If lngCnt > 0 Then
	Set rsUploads = Server.CreateObject("ADODB.RecordSet")
	strSQL = "SELECT [timestamp], [filename], [rid], [type], [uid] FROM uploads WHERE RID=" & strRID & " AND [staff]=0 ORDER BY [timestamp] DESC"
	rsUploads.Open strSQL, g_strCONNupload, 3, 1
	If rsUploads.EOF Then
		rsUploads.Close
		Set rsUploads = Nothing
		Session("MSG") = "Upload list cannot be accessed: " & strRID
		Response.Redirect "tsheet.asp?tmpDate=" & tmDate
		Response.Flush
		Response.End
	End If

	Do While Not rsUploads.EOF
		If ( rsUploads("type") = 0 ) Then
			strVform = strVform & "<li><a href=""dldup.asp?uid=" & rsUploads("uid") & """ download >" & _
					rsUploads("timestamp") & "</a></li>" & vbCrLf
			lngVFCnt = lngVFCnt + 1
		Else
			strTolls = strTolls & "<li><a href=""dldup.asp?uid=" & rsUploads("uid") & """ download >" & _
					rsUploads("timestamp") & "</a></li>" & vbCrLf
			lngTlCnt = lngTlCnt + 1
		End If
		If lngVFCnt > 0 Then strVForm = "<div class=""filelist""><h3>Verification Forms</h3><ul>" & strVForm & "</ul></div>"
		If lngTlCnt > 0 Then strTolls = "<div class=""filelist""><h3>Receipts</h3><ul>" & strTolls & "</ul></div>"

		rsUploads.MoveNext
	Loop

End If 'lngCnt > 0
%>
<!doctype html>
<html lang="en">

<head>
	<meta charset="utf-8">
	<meta http-equiv="x-ua-compatible" content="ie=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1">

	<title>Language Bank - Interpreter Uploads</title>
	<script language="JavaScript" type="text/JavaScript" src="js/jquery-3.3.1.min.js" ></script>
	<link href='style.css' type='text/css' rel='stylesheet' />
<style>
table.reqinfo {
	margin-left: 10%;
	width: 66.667%;
}
.reqinfo td:first-child {
	text-align: right;
	padding-right: 10px;
	vertical-align: top;
}
.reqinfo td {
	font-size: 10pt;
}
div.filelist {
	margin-right: 0px;
	width: 80%;
}
.filelist ol li {
	margin-left: 30px;
}
.filelist a, .filelist a:active, .filelist a:link, .filelist a:visited {
	color: #333;
	text-decoration: none;
}
.filelist a:hover {
	color: red;
	text-decoration: underline;
	background-color: yellow;
}
</style>
</head>
	<body>
<!-- #include file="_header.asp" -->
					<table cellSpacing='0' cellPadding='0' width="100%" border='0'>
							<!-- #include file="_greetme.asp" -->
						<tr>
						<td class='title' colspan='10' align='center'>File Uploads</td>
						</tr>
					</table>
<form action="vu_do.asp" method='post' id="frmUpload" name='frmUpload' enctype="multipart/form-data">
<table class="reqinfo">
	<thead></thead>
	<tbody>
		<tr><td>Request #:</td><td><%= rsTS("index")%>
				<input type="hidden" name="reqid" id="reqid" value="<%= rsTS("index")%>" readonly="readonly" />
					&nbsp;&nbsp;&nbsp;&nbsp;
					[<%= rsTS("statusname")%>]
					</td>

		<tr><td>Date:</td><td><%= rsTS("appdate")%>, <%= FormatDateTime(rsTS("AStarttime"), 4) %> to <%= FormatDateTime(rsTS("AEndtime"), 4) %>
				</td></tr>
		<tr><td>Activity:</td><td><b>
				<%= rsTS("Facility")%>
				<br />
				<%= rsTS("Dept")%>
				</b></td></tr>		
		<tr><td>Client:</td><td>
				<%= rsTS("cfname")%>&nbsp;&nbsp;<%= rsTS("clname")%></td></tr>
	</tbody>
</table>
<%
If lngCnt > 0 Then
	Response.Write "<div style=""width: 50%; margin-left: auto; margin-right: auto; padding-bottom: 20px;" & _
			" margin-top: 10px; border-top: 2px dotted #888; border-bottom: 2px dotted #888;"">"
	Response.Write "<p style=""font-size: 120%; font-weight: bold; line-height: 11pt;"">Files</p>"
	Response.Write strVForm
	Response.Write strTolls
	Response.Write "</div>"
End If
%>

<div style="width: 50%; margin-left: auto; margin-right: auto;">
	<h1>Upload Documents &mdash;&gt;&gt;</h1>
	<input type="hidden" name="rid" id="rid" value="<%= rsTS("index") %>" />
		<label for="ufile"><strong>Choose a file</strong></label>
		<input type="file" name="ufile" id="ufile" />
		<br />
		<table><tbody>
						<tr><td>
						<label for="utype"><strong>Upload type:</strong></label></td>
						<td>
						<input type="radio" value="0" name="utype" id="type_v" checked="checked" />&nbsp;Verification Form<br />
						<input type="radio" value="1" name="utype" id="type_t" />&nbsp;Toll and Parking Receipt<br />
						</td></tr>
					</tbody></table>
<div class="formatsmall">* PDF format preferred (JPG, PNG, GIF accepted)</div>
<div class="formatsmall">* 5MB file size limit</div>
<div class="formatsmall">* You can upload more than once, it will not overwrite previous upload</div>

<input type="submit" class="btn" name="btnUpld" id="btnUpld" value="Upload File" />
&nbsp;&nbsp;&nbsp;
<input type="button" class="btn" name="btnDn" id="btnDn" value="Back" />
</div>
</form>
<br /><br /><br /><br />

<table cellSpacing='0' cellPadding='0' style="width:100%;" border='0' class='bgstyle2'>
	<tr><td valign='bottom'><!-- #include file="_footer.asp" --></td></tr>
</table>
</body>
</html>
<%
rsTS.Close
Set rsTS = Nothing
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>								
<script>
$( document ).ready( function() {
		$('#btnDn').click(function(){
				console.log("aborting");
				document.location = "tsheet.asp?tmpDate=<%=tmDate%>";
			});
	});
</script>