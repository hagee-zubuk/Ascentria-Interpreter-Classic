<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- in file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
Function Z_MakeUniqueFileName()
	tmpdate = replace(date, "/", "") 
	tmpTime = replace(FormatDateTime(time, 3), ":", "")
	tmpTime = replace(tmpTime, " ", "")
	Z_MakeUniqueFileName = tmpdate & tmpTime
End Function
uploaderror = 0

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	Set oUpload = Server.CreateObject("SCUpload.Upload")
	oUpload.Upload
	If oUpload.Files.Count = 0 Then
		Set oUpload = Nothing
		Session("MSG") = "Please specify a file to import."
		Response.Redirect "upload.asp"
	End If
	appID = oUpload.Form("reqID")
	'create folder
	folderpath = uploadpath & appID
	folderpathvform = uploadpath & appID & "\vform"
	folderpathtoll = uploadpath & appID & "\tolls"
	Set fso = Server.CreateObject("Scripting.FileSystemObject")

	ServerShare = Z_UNFixPath(uploadpath)
	UserName = "acadatasrv2\LB_Webserv"
	Password = "1LBVerifyIdentity1"
	Set NetworkObject = CreateObject("WScript.Network")
On Error Resume Next
	NetworkObject.MapNetworkDrive "", ServerShare, False, UserName, Password
	If Err.Number<>0 Then
		Response.Write "Mapping: " & ServerShare & "<br />"
		Response.Write "<p>Error connecting to file repository!</p><pre>" & Err.Description & "</pre><br />Please contact LanguageBank<br />"
		Response.End
	End If
On Error Goto 0

	If Not fso.FolderExists(folderpath) Then fso.CreateFolder(folderpath)
	If Not fso.FolderExists(folderpathvform) Then fso.CreateFolder(folderpathvform)
	If Not fso.FolderExists(folderpathtoll) Then fso.CreateFolder(folderpathtoll)
	
	ctr = 1
	Do Until ctr = 3
		If ctr = 1 Then upfile = "Verification Form"
		If ctr = 2 Then upfile = "Tolls and Parking Receipts"
		oFileName = oUpload.Files(ctr).Item(1).filename
		If oFileName <> "" Then
			If Ucase(Z_GetExt(oFileName)) <> "PDF" Then
				Session("MSG") = Session("MSG") & "<br>" & upfile & " is invalid."
				uploaderror = 1
			End If
			oFileSize = oUpload.Files(ctr).Item(1).Size
			If oFileSize > 9000000 Then
				Session("MSG") = Session("MSG") & "<br>" & upfile & " is too large."
				uploaderror = 1
			End If
		Else
			uploaderror = 1
		End If
		If uploaderror = 0 Then
			'save file
			UniqueFilename = Z_MakeUniqueFileName()
			If ctr = 1 Then 
				nFileName = "vform" & UniqueFilename & ".PDF"
				folder = folderpathvform
				'oUpload.Files(ctr).Item(1).Save folderpathvform, nFileName
			ElseIf ctr = 2 Then 
				nFileName = "tollsandpark" & UniqueFilename & ".PDF"
				folder = folderpathtoll
			End If
			oUpload.Files(ctr).Item(1).Save "C:\WORK\LSS-LBIS\uploads", nfilename
			srcF = "C:\WORK\LSS-LBIS\uploads\" & nfilename
			fso.CopyFile srcF, Z_FixPath(folder), True
			Session("MSG") = "File Saved."

			'save in DB
			Set rsUpload = Server.CreateObject("ADODB.RecordSet")
			rsUpload.Open "SELECT * FROM uploads WHERE timestamp = '" & Now & "'", g_strCONNupload, 1, 3
			rsUpload.AddNew
			rsUpload("RID") = appID
			rsUpload("filename") = nFileName
			rsUpload("timestamp") = Now
			rsUpload("type") = False
			If ctr = 2 Then rsUpload("type") = True
			rsUpload.Update
			rsUpload.Close
			Set rsUpload = Nothing

			fso.DeleteFile srcF, True
		End If
		ctr = ctr + 1
	Loop
	Set oUpload = Nothing

	NetworkObject.RemoveNetworkDrive ServerShare, True, False
	Set NetworkObject = Nothing
	Set fso = Nothing
Else
	appID = Request("ReqID")
End If
%>
<html>
	<head>
		<title>Language Bank - Upload Documents</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function uploadFile(xxx) {
			if (document.frmUpload.F1.value == "" && document.frmUpload.F2.value == "") {
				alert("ERROR: Please select a file.")
				return;
			}
			if (document.frmUpload.F1.value != "") {
				filestr = document.frmUpload.F1.value.toUpperCase();
				if (filestr.indexOf(".PDF") == -1) {
					alert("ERROR: Incorrect file extension (Verification Form).")
					document.frmUpload.F1.value = "";
					return;
				}
			}
			if (document.frmUpload.F2.value != "") {
				filestr = document.frmUpload.F2.value.toUpperCase();
				if (filestr.indexOf(".PDF") == -1) {
					alert("ERROR: Incorrect file extension (Tolls and Parking Receipts.")
					document.frmUpload.F2.value = "";
					return;
				}
			}
			document.frmUpload.action = "upload.asp?reqID=" + xxx;
			document.frmUpload.submit();
		}
		-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form method='post' name='frmUpload' enctype="multipart/form-data">
			<table cellpadding='0' cellspacing='0' border='0' align='left' height='95%' width='100%'>
				<tr>
					<td class='header' colspan='2'><nobr>Upload Documents  --&gt&gt</td>
				</tr>
				<tr>
					<td align='right'>
						Verification Form:
					</td>
					<td>
							<input  class='main' type="file" name="F1" size="30" class='btn'>
					</td>
				</tr>
				<tr>
					<td align='right'>
						Tolls and Parking Receipt/s:
					</td>
					<td>
							<input  class='main' type="file" name="F2" size="30" class='btn' disabled>
					</td>
				</tr>
				<tr>
					<td colspan='2' align="center">
						
						<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*PDF format only</span><br>
						
						<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*5 MB limit per file</span><br>
						
						<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*You can upload more than once, it will not overwrite the previous upload</span>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td colspan='2' align='center'>
						<input class='btn' type='button' value='Upload' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='uploadFile(<%=appID%>);'>
							<input type="button" value="Close" class="btn" onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="self.close();">
						<input  type='hidden' name="reqID" value='<%=appID%>'>
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