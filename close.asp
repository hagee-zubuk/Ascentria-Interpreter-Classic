<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
%>
<html>
	<head>
		<title>Language Bank - Download Verification Form</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function dloadme(xxx)
		{
				document.frmVer.action = 'dl.asp?reqid=' + xxx;
				document.frmVer.submit();
		}
		-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form method='post' name='frmVer'>
			<table cellpadding='0' cellspacing='0' border='0' align='left' height='95%' width='100%'>
				<tr>
					<td class='header' colspan='2'><nobr>Download Verification Form  --&gt&gt</td>
				</tr>
				<tr>
					<td colspan='2' align='center'>
						<input class='btn' type='button' value='Download File' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='dloadme(<%=Request("reqid")%>);'>
						<input class='btn' type='button' value='Close' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='self.close();'>
					</td>
				</tr>
				
				<tr>
					<td colspan='3' align='left' valign='bottom'>
						<font size='1'><i>* Your computer/smart device must be able to read PDF files.</i></font><br>
						<font size='1'><i>* Download <a href='https://get.adobe.com/reader/' target='_blank'>HERE</a> to get a free PDF viewer.</i></font><br>
						<font size='1'><i>* Please close this window after downloading the file.</i></font>
					</td>
				</tr>		
			</table>
		</form>
	</body>
</html>
<%
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
