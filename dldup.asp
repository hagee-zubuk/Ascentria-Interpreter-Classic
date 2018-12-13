<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
lngUID = Z_CLng(Request("uid"))
Set rsUploads = Server.CreateObject("ADODB.RecordSet")
strSQL = "SELECT [filename], [type], [rid], [uid] FROM [uploads] WHERE UID=" & lngUID & " AND [staff]=0 ORDER BY [timestamp] DESC"
rsUploads.Open strSQL, g_strCONNupload, 3, 1
If rsUploads.EOF Then
	rsUploads.Close
	Set rsUploads = Nothing
	Response.Write "<h1>Oops</h1><p>Unable to find the file, or access is denied accessing the file.</p>"	
	'Response.Write strSQL
	Response.End
End If

strRID = Z_CLng(rsUploads("rid"))
If lngUID = rsUploads("uid") Then
	If ( rsUploads("type") = 0 ) Then
		subfold = "\vform\"
		file_nm = "Verification_Form"
	ElseIf ( rsUploads("type") = 1 ) Then
		subfold = "\tolls\"
		file_nm = "Receipt_Toll_Park"
	End If
	viewpath = uploadpath & strRID & subfold & rsUploads("filename")
End If
rsUploads.Close
Set rsUploads = Nothing

Set fso = CreateObject("Scripting.FileSystemObject")
strPath = Request("fpath")
strExt = LCase(Z_GetExt(viewpath))
If fso.FileExists(viewpath) Then	
	Set oFileStream = Server.CreateObject("ADODB.Stream")
	oFileStream.Open
	oFileStream.Type = 1 'Binary
On Error Resume Next
	oFileStream.LoadFromFile viewpath
	If Err.Number <> 0 Then
		Response.Write "Error: " & Err.Number & " [" & Err.Message & "]"
		Response.End
	End If
	Response.Clear
	If (strExt = "pdf") Then
		Response.ContentType = "application/pdf"
		Response.AddHeader "Content-Disposition", "inline; filename=v-form.pdf"
	Else
		Response.ContentType = "image/" & strExt
		Response.AddHeader "Content-Disposition", "inline; filename=v-form." & strExt
	End If

	Dim lSize, lBlocks
	'Const CHUNK = 2048
	Const CHUNK = 20480
	lSize = oFileStream.Size
	Response.AddHeader "Content-Size", lSize
	lBlocks = 1
	Response.Buffer = False
	Do Until oFileStream.EOS Or Not Response.IsClientConnected
		Response.BinaryWrite(oFileStream.Read(CHUNK))
	Loop
	'Response.BinaryWrite oFileStream.Read
	oFileStream.Close
	Set oFileStream= Nothing
Else
	Response.Write "<h1>Oops</h1><p>Unable to find the file, or access is denied accessing the file.</p>"
	'Response.Write viewpath
End If
Set fso = Nothing
%>

